VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMediLists 
   BackColor       =   &H8000000C&
   Caption         =   "药品目录管理"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   11205
   Icon            =   "frmMediLists.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picHBar 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   2805
      MousePointer    =   7  'Size N S
      ScaleHeight     =   30
      ScaleWidth      =   6075
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3870
      Width           =   6075
   End
   Begin VB.PictureBox picVBar 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6660
      Left            =   2580
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6660
      ScaleWidth      =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   795
      Width           =   30
   End
   Begin VB.PictureBox picClass 
      Height          =   6735
      Left            =   0
      ScaleHeight     =   6675
      ScaleWidth      =   2340
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   2400
      Begin VB.CommandButton cmdKind 
         Caption         =   "过滤结果"
         Height          =   300
         Index           =   4
         Left            =   0
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   1155
         Width           =   2295
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "特性药品"
         Height          =   300
         Index           =   3
         Left            =   0
         TabIndex        =   23
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   870
         Width           =   2295
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "中草药(&7)"
         Height          =   300
         Index           =   2
         Left            =   0
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   585
         Width           =   2295
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "中成药(&6)"
         Height          =   300
         Index           =   1
         Left            =   0
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   300
         Width           =   2295
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "西成药(&5)"
         Height          =   300
         Index           =   0
         Left            =   0
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   15
         Width           =   2295
      End
      Begin MSComctlLib.TreeView tvwClass 
         Height          =   4800
         Left            =   0
         TabIndex        =   5
         Tag             =   "1000"
         Top             =   1440
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   8467
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "imgList"
         Appearance      =   0
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1785
      Top             =   6750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":030A
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":08A4
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":0E3E
            Key             =   "成药U"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":13D8
            Key             =   "成药S"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":1972
            Key             =   "规格U"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":1F0C
            Key             =   "规格S"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":24A6
            Key             =   "草药U"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":2A40
            Key             =   "草药S"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":2FDA
            Key             =   "Packer"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":3574
            Key             =   "NoPacker"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":3B0E
            Key             =   "草规S"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":43E8
            Key             =   "草规U"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2895
      Left            =   2835
      TabIndex        =   1
      Top             =   855
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   7935
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMediLists.frx":4CC2
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14684
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
   Begin ComCtl3.CoolBar clbThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   11205
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbThis"
      MinWidth1       =   24000
      MinHeight1      =   720
      Width1          =   8730
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   720
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   24000
         _ExtentX        =   42333
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
            NumButtons      =   17
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Description     =   "预览"
               Object.ToolTipText     =   "预览当前表"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印当前表"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "分类"
               Key             =   "Class"
               Description     =   "分类"
               Object.ToolTipText     =   "调整药品分类"
               Object.Tag             =   "分类"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "品种"
               Key             =   "Item"
               Description     =   "品种"
               Object.ToolTipText     =   "调整药品品种"
               Object.Tag             =   "品种"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "规格"
               Key             =   "Spec"
               Description     =   "规格"
               Object.ToolTipText     =   "调整同种药品的规格"
               Object.Tag             =   "规格"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "sp2"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "启用"
               Key             =   "Start"
               Description     =   "启用"
               Object.ToolTipText     =   "启用指定的停用药品"
               Object.Tag             =   "启用"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "停用"
               Key             =   "Stop"
               Description     =   "停用"
               Object.ToolTipText     =   "停用指定的在用药品"
               Object.Tag             =   "停用"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "储备"
               Key             =   "Limit"
               Description     =   "储备"
               Object.ToolTipText     =   "调整储备限额"
               Object.Tag             =   "储备"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查找"
               Key             =   "Find"
               Description     =   "查找"
               Object.ToolTipText     =   "查找诊断条目"
               Object.Tag             =   "查找"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filter"
               Description     =   "过滤"
               Object.ToolTipText     =   "过滤药品目录"
               Object.Tag             =   "过滤"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
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
         Begin VB.PictureBox picFind 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   9000
            ScaleHeight     =   285.714
            ScaleMode       =   0  'User
            ScaleWidth      =   495
            TabIndex        =   26
            Top             =   210
            Width           =   495
            Begin VB.Label lbl查找 
               Caption         =   "查找"
               Height          =   255
               Left            =   120
               TabIndex        =   27
               Top             =   75
               Width           =   495
            End
         End
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   9600
            MaxLength       =   10
            TabIndex        =   25
            Tag             =   "简码"
            Top             =   240
            Width           =   1425
         End
      End
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   7680
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":5554
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":576E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":5988
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":5BA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":5DBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":64B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":66D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":68EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":6FE4
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":71FE
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":741E
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":763E
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":7958
            Key             =   "Limit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   6915
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":8052
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":8272
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":8492
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":86AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":88C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":8FC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":91DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":93F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":9AEE
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":9D08
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":9F28
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":A148
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLists.frx":A462
            Key             =   "Limit"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabContent 
      Height          =   2805
      HelpContextID   =   1
      Left            =   2760
      TabIndex        =   10
      Top             =   4635
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   4948
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "药品规格(&S)"
      TabPicture(0)   =   "frmMediLists.frx":AB5C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraComment(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lvwSpecs"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "售价记录(&L)"
      TabPicture(1)   =   "frmMediLists.frx":AB78
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "hgdPrice"
      Tab(1).Control(1)=   "fraComment(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "成本价记录(&C)"
      TabPicture(2)   =   "frmMediLists.frx":AB94
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "hgdCost"
      Tab(2).Control(1)=   "chkStock"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "费别等级(&F)"
      TabPicture(3)   =   "frmMediLists.frx":ABB0
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "hgdCharge"
      Tab(3).ControlCount=   1
      Begin VB.CheckBox chkStock 
         Caption         =   "只显示有库存价格记录"
         Height          =   180
         Left            =   -69840
         TabIndex        =   28
         Top             =   80
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin MSComctlLib.ListView lvwSpecs 
         Height          =   1395
         Left            =   105
         TabIndex        =   12
         Top             =   405
         Width           =   7410
         _ExtentX        =   13070
         _ExtentY        =   2461
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgList"
         SmallIcons      =   "imgList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdPrice 
         Height          =   1665
         Left            =   -74880
         TabIndex        =   11
         Top             =   360
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   2937
         _Version        =   393216
         Rows            =   4
         Cols            =   13
         FixedCols       =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   13
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdCost 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   21
         Top             =   360
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3836
         _Version        =   393216
         Rows            =   4
         Cols            =   9
         FixedCols       =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   9
      End
      Begin VB.Frame fraComment 
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   0
         Left            =   90
         TabIndex        =   13
         Top             =   1860
         Width           =   7410
         Begin VB.Label lblComment 
            AutoSize        =   -1  'True
            Caption         =   "3、分零后效用克保持3天"
            Height          =   180
            Index           =   2
            Left            =   0
            TabIndex        =   16
            Top             =   510
            Width           =   1980
         End
         Begin VB.Label lblComment 
            AutoSize        =   -1  'True
            Caption         =   "2、不进行分批与效期跟踪管理"
            Height          =   180
            Index           =   1
            Left            =   0
            TabIndex        =   15
            Top             =   270
            Width           =   2430
         End
         Begin VB.Label lblComment 
            AutoSize        =   -1  'True
            Caption         =   "1、售价单位：片，门诊单位：瓶"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   14
            Top             =   30
            Width           =   2610
         End
      End
      Begin VB.Frame fraComment 
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   1
         Left            =   -74910
         TabIndex        =   17
         Top             =   2085
         Width           =   7410
         Begin VB.Label lblComment 
            AutoSize        =   -1  'True
            Caption         =   "1、时价药，指导批发价6元/瓶。。。"
            Height          =   180
            Index           =   3
            Left            =   0
            TabIndex        =   19
            Top             =   30
            Width           =   2970
         End
         Begin VB.Label lblComment 
            AutoSize        =   -1  'True
            Caption         =   "2、最高售价198.25元/瓶；根据病人身份费别进行优惠或加价。"
            Height          =   180
            Index           =   4
            Left            =   0
            TabIndex        =   18
            Top             =   270
            Width           =   5040
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdCharge 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   22
         Top             =   360
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3836
         _Version        =   393216
         Rows            =   4
         Cols            =   9
         FixedCols       =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   9
      End
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
      Begin VB.Menu mnuFileSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePara 
         Caption         =   "参数设置(&A)"
         Shortcut        =   {F12}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSpt2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuClass 
      Caption         =   "分类(&K)"
      Begin VB.Menu mnuClassAdd 
         Caption         =   "新增(&I)"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu mnuClassMod 
         Caption         =   "修改(&U)"
      End
      Begin VB.Menu mnuClassDel 
         Caption         =   "删除(&E)"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuClassSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClassStar 
         Caption         =   "启用分类(&R)"
      End
      Begin VB.Menu mnuClassStop 
         Caption         =   "停用分类(&S)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "药品(&E)"
      Begin VB.Menu mnuEditItemAdd 
         Caption         =   "新增品种(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditItemMod 
         Caption         =   "修改品种(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditItemDel 
         Caption         =   "删除品种(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSpt6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditItemTabu 
         Caption         =   "配伍禁忌(&T)..."
      End
      Begin VB.Menu mnuEditItemUsage 
         Caption         =   "用法用量(&U)"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuEditItemBill 
         Caption         =   "对应处方(&B)"
      End
      Begin VB.Menu mnuEditSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSpecAdd 
         Caption         =   "新增规格(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuEditSpecMod 
         Caption         =   "修改规格(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuEditSpecDel 
         Caption         =   "删除规格(&Y)"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEditSpecExp 
         Caption         =   "规格扩展信息定义(&E)"
      End
      Begin VB.Menu mnuEditSpt7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditItemPart 
         Caption         =   "存储库房(&R)..."
      End
      Begin VB.Menu mnuEditSpecLimit 
         Caption         =   "储备限量(&L)..."
      End
      Begin VB.Menu mnuEditSpecProtocol 
         Caption         =   "协定药品(&P)..."
      End
      Begin VB.Menu mnuEditSpecSelf 
         Caption         =   "自制药品(&H)..."
      End
      Begin VB.Menu mnuEditSpecUnit 
         Caption         =   "中标单位(&V)..."
      End
      Begin VB.Menu mnuEditManFac 
         Caption         =   "厂家批准文号(&C)"
      End
      Begin VB.Menu mnuEditSendType 
         Caption         =   "发药类型(&S)..."
      End
      Begin VB.Menu mnuEditSpt5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditRate 
         Caption         =   "分段加成率(&L)"
      End
      Begin VB.Menu mnuEditSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditVariBatch 
         Caption         =   "品种批量修改(&W)"
      End
      Begin VB.Menu mnuEditSpecBatch 
         Caption         =   "规格批量修改(&Y)"
      End
      Begin VB.Menu mnuEditExcel 
         Caption         =   "导入项目"
      End
      Begin VB.Menu mnuEditContrast 
         Caption         =   "带量采购对照(&G)"
      End
      Begin VB.Menu mnuEditSpt3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStart 
         Caption         =   "启用(&R)"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "停用(&S)"
      End
      Begin VB.Menu mnuEditSpt4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPriceChargeSet1 
         Caption         =   "费别设置(&C)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSptPacker 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUploadDrugInfo 
         Caption         =   "批量上传物流平台"
      End
   End
   Begin VB.Menu mnuPrice 
      Caption         =   "价格(&P)"
      Begin VB.Menu mnuPriceChargeSet 
         Caption         =   "费别设置(&C)"
      End
      Begin VB.Menu mnuPriceSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPriceTable 
         Caption         =   "调价记录表(&S)"
      End
      Begin VB.Menu mnuPriceLists 
         Caption         =   "药品价目表(&L)..."
         Shortcut        =   ^L
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
      Begin VB.Menu mnuViewStates 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShowAll 
         Caption         =   "显示所有下级(&L)"
      End
      Begin VB.Menu mnuViewList 
         Caption         =   "显示停用目录(&M)"
      End
      Begin VB.Menu mnuViewStoped 
         Caption         =   "显示停用药品(&C)"
      End
      Begin VB.Menu mnuViewPrices 
         Caption         =   "显示历史价格(&H)"
      End
      Begin VB.Menu mnuViewSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "查找(&F)..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewFindNext 
         Caption         =   "查找下一条(&N)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "过滤(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuViewRefer 
         Caption         =   "参考(&R)..."
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuViewSpt3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web上的中联(&W)"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&E)..."
         End
      End
      Begin VB.Menu mnuHelp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmMediLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngMode As Long
Public mstrPrivs As String       '用户具有本程序的具体权限

Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim intCount As Integer, intRow As Integer, intCol As Integer
Dim strTemp As String
Dim mint药库单位  As Integer
Dim mstrType As String              '过滤窗口返回的药品材质类型串  '5-西成药 6-中成药 7-中草药
Dim mstrDrugId As String            '过滤窗口返回的药名ID串
Dim mlngCurrDrug As Long
Private mstrDBNodeClick As String
Private mstrNodeClick西成药 As String     '记录上次选中的分类
Private mstrNodeClick中成药 As String
Private mstrNodeClick中草药 As String
Private mstrNodeSelect西成药 As String
Private mstrNodeSelect中成药 As String
Private mstrNodeSelect中草药 As String
Private mstrItemClick西成药 As String     '记录上次选中的药品
Private mstrItemClick中成药 As String     '记录上次选中的药品
Private mstrItemClick中草药 As String     '记录上次选中的药品
Private mstrKey As String           '记录所所选中的分类
Private Declare Function SetParent Lib "user32 " (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private mstrFindValue As String     '查找字符串
Private mrsFind As ADODB.Recordset  '记录查询的数据集
Private mstr分类 As String          '记录是选择的什么类型的药品 "0"-西成药,"1"-中成药,"2"-中草药
Private mintPage As Integer         '记录当前所选中的SSTAB页面
Private mStrItem As String          '选中的品种节点

Private Const colPrice药价类型 As Integer = 0
Private Const colPrice指导批价 As Integer = 1
Private Const colPrice扣率 As Integer = 2
Private Const colPrice指导售价 As Integer = 3
Private Const colPrice指导差率 As Integer = 4
Private Const colPrice屏蔽费别 As Integer = 5
Private Const colPrice执行情况 As Integer = 6
Private Const colPrice调价NO As Integer = 7
Private Const colPrice药品 As Integer = 8
Private Const colPrice单位 As Integer = 9
Private Const colPrice售价 As Integer = 10
Private Const colPrice收入项目 As Integer = 11
Private Const colPrice执行日期 As Integer = 12
Private Const colPrice说明 As Integer = 13
Private Const colPrice药品ID As Integer = 14

Private Const colCost药品id As Integer = 0
Private Const colCostNO As Integer = 1
Private Const colCost药品 As Integer = 2
Private Const colCost库房 As Integer = 3

Private Const colCost批号 As Integer = 4
Private Const colCost效期 As Integer = 5
Private Const colCost产地 As Integer = 6
Private Const colCost单位 As Integer = 7
Private Const colCost原成本价 As Integer = 8
Private Const colCost成本价 As Integer = 9
Private Const colCost执行日期 As Integer = 10
Private Const colCost说明 As Integer = 11

Private Const col药品 As Integer = 0
Private Const col库房 As Integer = 1
Private Const col上限 As Integer = 2
Private Const col下限 As Integer = 3
Private Const col日盘 As Integer = 4
Private Const col周盘 As Integer = 5
Private Const col月盘 As Integer = 6
Private Const col季盘 As Integer = 7
Private Const col货位 As Integer = 8

Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mintCostDigit As Integer        '成本价小数位数
Private mintPriceDigit As Integer       '售价小数位数
Private mbln自管药 As Boolean           '用来记录是否是通过自管药设置方式打开的窗体
Private mintIndex As Integer    '用来记录被点击的分类

Private Const mconColor_Stop As Long = &HFF&

Public Sub ShowMe(ByVal frmPar As Form, ByVal bln自管药 As Boolean)
    '显示窗体
    mbln自管药 = bln自管药
    Me.Show , frmPar
End Sub

Private Sub GetCostAdjust(ByVal lngDrug As Long)
    Dim strSqlCon As String
    
    '----------填写成本价-----------------
    On Error GoTo ErrHandle
    With Me.hgdCost
        .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
        For intCol = 0 To .Cols - 1
            .TextMatrix(.FixedRows, intCol) = ""
        Next
    End With
    
    If Me.mnuViewStoped.Checked = False Then
        strSqlCon = " and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    End If
    
    gstrSql = " Select B.NO, I.ID As 药品id, '[' || I.编码 || ']' || I.名称 || ' ' || I.规格 || ' ' || I.产地 As 药品, P.名称 As 库房,A.批号,A.效期,A.产地, " & _
            " I.计算单位 As 单位, S.药库单位, Nvl(S.药库包装, 1) 药库包装, A.原价 As 原成本价,A.现价 As 成本价, A.执行日期, A.调价说明 " & _
            " From 药品收发记录 B, 收费项目目录 I, 药品规格 S, 部门表 P, 药品价格记录 A " & _
            " Where A.价格类型=2 And A.收发id = B.ID(+) And A.药品id = I.ID And " & _
            " I.ID = S.药品id And A.库房id = P.ID(+) And S.药名id = [1] " & strSqlCon
    
    If chkStock.Value = 1 Then
        gstrSql = gstrSql & " And Exists (Select 1 From 药品库存 K Where 性质 = 1 And k.库房id = a.库房id And k.药品id = a.药品id And k.批次 = a.批次) "
    End If
    
    gstrSql = gstrSql & " Order By 药品, 执行日期 Desc, NO Desc "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngDrug)
    
    With rsTemp
        Me.hgdCost.Redraw = False
        If .BOF Or .EOF Then
            With Me.hgdCost
                .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
                For intCol = 0 To .Cols - 1
                    .TextMatrix(.FixedRows, intCol) = ""
                Next
            End With
        Else
            Me.hgdCost.Rows = Me.hgdCost.FixedRows + .RecordCount
        End If
        Do While Not .EOF
            Me.hgdCost.RowData(.AbsolutePosition) = !药品id
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCostNO) = IIf(IsNull(!No), "", !No)
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCost药品) = !药品
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCost库房) = IIf(IsNull(!库房), "", !库房)
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCost批号) = IIf(IsNull(!批号), "", !批号)
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCost效期) = IIf(IsNull(!效期), "", Format(!效期, "yyyy-mm-dd"))
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCost产地) = IIf(IsNull(!产地), "", !产地)
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCost单位) = IIf(mint药库单位 = 0, !单位, !药库单位)
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCost原成本价) = Format(!原成本价 * IIf(mint药库单位 = 0, 1, !药库包装), mstrCostFormat)
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCost成本价) = Format(!成本价 * IIf(mint药库单位 = 0, 1, !药库包装), mstrCostFormat)
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCost执行日期) = IIf(IsNull(!执行日期), "", Format(!执行日期, "yyyy-mm-dd hh:mm:ss"))
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCost说明) = IIf(IsNull(!调价说明), "", !调价说明)
            Me.hgdCost.TextMatrix(.AbsolutePosition, colCost药品id) = !药品id
            
            Me.hgdCost.Row = .AbsolutePosition
            For intCol = 0 To Me.hgdCost.Cols - 1
                Me.hgdCost.Col = intCol
                If IIf(IsNull(!执行日期), "", !执行日期) = "" Then
                    Me.hgdCost.CellBackColor = RGB(225, 255, 255)
                Else
                    Me.hgdCost.CellBackColor = RGB(240, 240, 240)
                End If
            Next
            .MoveNext
        Loop
        Me.hgdCost.Row = Me.hgdCost.FixedRows

        Me.hgdCost.Redraw = True
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetChargeSet(ByVal lngDrug As Long)
    Dim strSqlCon As String
    
    '----------填写成本价-----------------
    On Error GoTo ErrHandle
    With Me.hgdCharge
        .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
        For intCol = 0 To .Cols - 1
            .TextMatrix(.FixedRows, intCol) = ""
        Next
    End With
    
    gstrSql = "Select B.ID, '[' || B.编码 || ']' || B.名称 As 名称, A.费别, A.段号, 应收段首值, 应收段尾值, 实收比率, Decode(计算方法, 1, '1-成本价加收比例计算', '0-分段比例计算') As 计算方法 " & _
        " From 费别明细 A, 收费项目目录 B, 药品规格 C " & _
        " Where A.收费细目id = B.ID And B.ID = C.药品id And C.药名id = [1] " & _
        " Order By 名称, A.费别, A.段号, A.应收段首值"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngDrug)
    
    With rsTemp
        Me.hgdCharge.Redraw = False
        If .BOF Or .EOF Then
            With Me.hgdCharge
                .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
                For intCol = 0 To .Cols - 1
                    .TextMatrix(.FixedRows, intCol) = ""
                Next
            End With
        Else
            Me.hgdCharge.Rows = Me.hgdCharge.FixedRows + .RecordCount
        End If
        Do While Not .EOF
            Me.hgdCharge.RowData(.AbsolutePosition) = !ID
            Me.hgdCharge.TextMatrix(.AbsolutePosition, 0) = .Fields("名称").Value
            Me.hgdCharge.TextMatrix(.AbsolutePosition, 1) = .Fields("费别").Value
            Me.hgdCharge.TextMatrix(.AbsolutePosition, 2) = Format(.Fields("应收段首值").Value, "##########0.00;-#########0.00;0.00;0.00") & _
                " ～ " & Format(.Fields("应收段尾值").Value, "##########0.00;-#########0.00;0.00;0.00")
            Me.hgdCharge.TextMatrix(.AbsolutePosition, 3) = Format(.Fields("实收比率").Value, "###0.00;-##0.00;0.00;0.00")
            Me.hgdCharge.TextMatrix(.AbsolutePosition, 4) = .Fields("计算方法").Value
            
            .MoveNext
        Loop
        
        Me.hgdCharge.ColAlignment(2) = 1
        Me.hgdCharge.ColAlignment(3) = 1
        Me.hgdCharge.MergeCells = flexMergeRestrictColumns
        Me.hgdCharge.MergeCol(0) = True
        Me.hgdCharge.MergeCol(1) = True
        
        Me.hgdCharge.Redraw = True
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Sub zlGetFilter(ByVal strType As String, ByVal strDrugId As String)
    mstrType = strType
    mstrDrugId = strDrugId
    Call cmdKind_Click(4)
End Sub

Private Sub zlPopupClassMenu()
    With tvwClass
        Select Case .Tag
        Case 0
            If InStr(1, mstrPrivs, ";管理西成药;") = 0 Then
                mnuClassAdd.Visible = False
                mnuClassMod.Visible = False
                mnuClassDel.Visible = False
            Else
                mnuClassAdd.Visible = True
                mnuClassMod.Visible = True
                mnuClassDel.Visible = True
            End If
        Case 1
            If InStr(1, mstrPrivs, ";管理中成药;") = 0 Then
                mnuClassAdd.Visible = False
                mnuClassMod.Visible = False
                mnuClassDel.Visible = False
            Else
                mnuClassAdd.Visible = True
                mnuClassMod.Visible = True
                mnuClassDel.Visible = True
            End If
        Case 2
            If InStr(1, mstrPrivs, ";管理中草药;") = 0 Then
                mnuClassAdd.Visible = False
                mnuClassMod.Visible = False
                mnuClassDel.Visible = False
            Else
                mnuClassAdd.Visible = True
                mnuClassMod.Visible = True
                mnuClassDel.Visible = True
            End If
        End Select
    End With
    If InStr(1, mstrPrivs, ";药品启用;") = 0 Then
        mnuClassStar.Visible = False
    Else
        mnuClassStar.Visible = True
    End If
    If InStr(1, mstrPrivs, ";药品停用;") = 0 Then
        mnuClassStop.Visible = False
    Else
        mnuClassStop.Visible = True
    End If
                
    If mnuClassAdd.Visible = False And mnuClassMod.Visible = False And mnuClassDel.Visible = False And mnuClassStar.Visible = False And mnuClassStop.Visible = False Then
        Exit Sub
    End If
    Set objNode = Me.tvwClass.SelectedItem
    
    If objNode Is Nothing Then
        Me.mnuClassMod.Enabled = False
        Me.mnuClassDel.Enabled = False
        Me.mnuClassStar.Enabled = False
        Me.mnuClassStop.Enabled = False
    Else
        If Val(objNode.Tag) <= 2 Then
            If objNode.ForeColor = mconColor_Stop Then
                Me.mnuClassAdd.Enabled = False
                Me.mnuClassMod.Enabled = False
                Me.mnuClassStar.Enabled = True
                Me.mnuClassStop.Enabled = False
            Else
                Me.mnuClassAdd.Enabled = True
                Me.mnuClassMod.Enabled = True
                Me.mnuClassStar.Enabled = False
                Me.mnuClassStop.Enabled = True
            End If
        Else
            Me.mnuClassStar.Enabled = False
            Me.mnuClassStop.Enabled = False
        End If
    End If
    
    Call setMenu自管药
    Call PopupMenu(Me.mnuClass, 2)
End Sub

Private Sub chkStock_Click()
    Call GetCostAdjust(mlngCurrDrug)
End Sub

Private Sub cmdKind_Click(Index As Integer)
    Dim intCount As Integer
    Dim objNode As Node
    Dim strTemp As String
    Dim strItem As String
    
    mintIndex = Index
    mstrKey = ""
    mstrFindValue = ""
    Call SaveListViewState(Me.lvwItems, Me.Name & Val(Me.tvwClass.Tag), Me.lvwItems.View)
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        If intCount <= Index Then
            Me.cmdKind(intCount).Tag = 0
        Else
            Me.cmdKind(intCount).Tag = 1
        End If
    Next
    
    mstr分类 = Index & ""
    '装数据并调整界面
    If Me.lvwItems.Visible Then
        Call picClass_Resize
        Me.tvwClass.SetFocus
    End If
    If Index < 3 Then
        If Val(tvwClass.Tag) <> Index Then
            Me.tvwClass.Tag = Index
            Call zlRefClasses
        End If
        Me.mnuViewFind.Enabled = True
        Me.mnuViewFindNext.Enabled = True
        Me.tlbThis.Buttons("Find").Enabled = True
        Me.mnuEditExcel.Enabled = True
    Else
        Me.tvwClass.Tag = Index
        Call zlRefClasses
        Me.mnuViewFind.Enabled = False
        Me.mnuViewFindNext.Enabled = False
        Me.tlbThis.Buttons("Find").Enabled = False
        Me.mnuEditExcel.Enabled = False
        frmMediFind.Hide
    End If
    If Val(tvwClass.Tag) >= 3 Then
        txtFind.Enabled = False
        txtFind.BackColor = &H8000000F  '灰色不可用
    Else
        txtFind.Enabled = True
        txtFind.BackColor = vbWhite '白色可以修改
    End If
    
    If mstr分类 = "0" Then
        strTemp = mstrNodeSelect西成药
    ElseIf mstr分类 = "1" Then
        strTemp = mstrNodeSelect中成药
    ElseIf mstr分类 = "2" Then
        strTemp = mstrNodeSelect中草药
    End If
    
    For Each objNode In tvwClass.Nodes
        If objNode.Key = strTemp Then
            objNode.Selected = True
            Call tvwClass_NodeClick(objNode)
            Exit For
        End If
    Next
End Sub

Private Sub clbThis_Resize()
    Me.clbThis.Bands(1).MinHeight = Me.tlbThis.Height
    Me.clbThis.Refresh
    Call Form_Resize
End Sub

Private Sub Form_Activate()
    Me.lvwItems.Visible = True
    Call Form_Resize
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    Dim lngCount As Long
    Dim rs收入项目 As ADODB.Recordset
    Dim bln收入项目 As Boolean
    
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    
    gblnIncomeItem = False
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    If mbln自管药 = True Then '自管药不调用模块中取输入匹配的值，自管药单独处理
        gstrMatch = IIf(Val(zlDatabase.GetPara("输入匹配", , , True)) = 0, "%", "")
    End If
    
    If GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", "1") = "1" Then
        strTemp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\分割", "横向", "0")
        If strTemp <> "0" Then
            Me.picVBar.Left = CLng(strTemp)
        End If
        strTemp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\分割", "纵向", "0")
        If strTemp <> "0" Then
            Me.picHBar.Top = CLng(strTemp)
        End If
    End If
    
    mnuViewShowAll.Checked = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示所有", 0)) = 1)
    mnuViewList.Checked = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示停用目录", 0)) = 1)
    mnuViewStoped.Checked = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示停用项目", 0)) = 1)
    mnuViewPrices.Checked = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示历史价格", 0)) = 1)
    
    '检查是否以药库单位显示价格
    mint药库单位 = Val(zlDatabase.GetPara(29, glngSys))
    
    mintCostDigit = GetDigit(1, 1, IIf(mint药库单位 = 0, 1, 4))
    mintPriceDigit = GetDigit(1, 2, IIf(mint药库单位 = 0, 1, 4))
    
    mstrCostFormat = "0." & String(mintCostDigit, "0") & ";-0." & String(mintCostDigit, "0") & ";0"
    mstrPriceFormat = "0." & String(mintPriceDigit, "0") & ";-0." & String(mintPriceDigit, "0") & ";0"
    
    
    '可直接通过菜单进行的权限控制
'    If InStr(1, mstrPrivs, "参数设置") = 0 Then Me.mnuFilePara.Visible = False: Me.mnuFileSpt2.Visible = False
    If InStr(1, mstrPrivs, "协定药品构成") = 0 Then Me.mnuEditSpecProtocol.Visible = False:
    
    If InStr(1, mstrPrivs, "药品停用") = 0 Then Me.mnuEditStop.Visible = False:
    If InStr(1, mstrPrivs, "药品启用") = 0 Then Me.mnuEditStart.Visible = False:
    If InStr(1, mstrPrivs, "批量设置发药类型") = 0 Then Me.mnuEditSendType.Visible = False
    Me.mnuEditSpt2.Visible = (Me.mnuEditStop.Visible Or Me.mnuEditStart.Visible)
    
    tlbThis.Buttons("Stop").Visible = Me.mnuEditStop.Visible
    tlbThis.Buttons("Start").Visible = Me.mnuEditStart.Visible
    tlbThis.Buttons(7).Visible = Me.mnuEditSpt2.Visible
    
    '售价表格设置
    With Me.hgdPrice
        .Redraw = False
        .Rows = .FixedRows + 1: .Cols = 15
        
        .TextMatrix(0, colPrice药品) = "药品": .TextMatrix(0, colPrice单位) = "单位": .TextMatrix(0, colPrice售价) = "售价"
        .TextMatrix(0, colPrice收入项目) = "收入项目": .TextMatrix(0, colPrice说明) = "说明": .TextMatrix(0, colPrice执行日期) = "执行日期"
        .TextMatrix(0, colPrice药品ID) = "药品ID"
        .TextMatrix(0, colPrice调价NO) = "调价单据号"
        
        .ColWidth(colPrice药价类型) = 0: .ColWidth(colPrice指导批价) = 0: .ColWidth(colPrice扣率) = 0: .ColWidth(colPrice指导售价) = 0
        .ColWidth(colPrice指导差率) = 0: .ColWidth(colPrice屏蔽费别) = 0: .ColWidth(colPrice执行情况) = 0
        .ColWidth(colPrice药品) = 3500: .ColWidth(colPrice单位) = 550: .ColWidth(colPrice售价) = 1000
        .ColWidth(colPrice收入项目) = 850: .ColWidth(colPrice说明) = 2500: .ColWidth(colPrice执行日期) = 1800
        .ColWidth(colPrice药品ID) = 0
        .ColWidth(colPrice调价NO) = 1000
        
        For intCol = 0 To .Cols - 1
            .ColAlignmentFixed(intCol) = 4
        Next
        .ColAlignment(colPrice药品) = 1: .ColAlignment(colPrice单位) = 4: .ColAlignment(colPrice售价) = 7
        .ColAlignment(colPrice收入项目) = 1: .ColAlignment(colPrice说明) = 1: .ColAlignment(colPrice执行日期) = 1
        .Redraw = True
    End With
    
    '成本价表格设置
    With Me.hgdCost
        .Redraw = False
        .Rows = .FixedRows + 1
        .Cols = 12
        
        .TextMatrix(0, colCost药品id) = "药品id"
        .TextMatrix(0, colCostNO) = "调价NO"
        .TextMatrix(0, colCost药品) = "药品"
        .TextMatrix(0, colCost库房) = "库房"
        .TextMatrix(0, colCost批号) = "批号"
        .TextMatrix(0, colCost效期) = "效期"
        .TextMatrix(0, colCost产地) = "产地"
        .TextMatrix(0, colCost单位) = "单位"
        .TextMatrix(0, colCost原成本价) = "原成本价"
        .TextMatrix(0, colCost成本价) = "新成本价"
        .TextMatrix(0, colCost执行日期) = "执行日期"
        .TextMatrix(0, colCost说明) = "说明"
        
        .ColWidth(colCost药品id) = 0
        .ColWidth(colCostNO) = 1000
        .ColWidth(colCost药品) = 3500
        .ColWidth(colCost库房) = 1500
        .ColWidth(colCost批号) = 1000
        .ColWidth(colCost效期) = 1000
        .ColWidth(colCost产地) = 1000
        .ColWidth(colCost单位) = 550
        .ColWidth(colCost原成本价) = 1000
        .ColWidth(colCost成本价) = 1000
        .ColWidth(colCost执行日期) = 1800
        .ColWidth(colCost说明) = 2500
        
        For intCol = 0 To .Cols - 1
            .ColAlignmentFixed(intCol) = 4
        Next
        
        .ColAlignment(colCostNO) = 1
        .ColAlignment(colCost药品) = 1
        .ColAlignment(colCost库房) = 1
        .ColAlignment(colCost批号) = 1
        .ColAlignment(colCost效期) = 1
        .ColAlignment(colCost产地) = 1
        .ColAlignment(colCost单位) = 4
        .ColAlignment(colCost原成本价) = 7
        .ColAlignment(colCost成本价) = 7
        .ColAlignment(colCost执行日期) = 1
        .ColAlignment(colCost说明) = 1
        
        .Redraw = True
    End With
    
    '费别等级列表设置
    With hgdCharge
        .Cols = 5
        .ColWidth(0) = 4000
        .ColWidth(1) = 1500
        .ColWidth(2) = 3000
        .ColWidth(3) = 1050
        .ColWidth(4) = 2000
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .TextMatrix(0, 0) = "药品"
        .TextMatrix(0, 1) = "费别"
        .TextMatrix(0, 2) = "应收金额(元)"
        .TextMatrix(0, 3) = "实收比率(%)"
        .TextMatrix(0, 4) = "计算方法"
        
        .MergeCol(0) = True
'        .MergeCol(1) = True
    End With
    
    Me.picHBar.Top = Me.ScaleHeight - IIf(Me.stbThis.Visible, Me.stbThis.Height, 0) - 2500
    Call cmdKind_Click(0)
    
    If mbln自管药 = False Then
        '物流平台接口
        On Error Resume Next
        LogisticPlatformInterface
    End If
    
    '是否启用抗菌药物严格控制
    gblnKSSStrict = CheckKSSPrivilege
    
    If gblnKSSStrict = False Then
        lngCount = 0
        For i = 0 To Me.mnuReportItem.UBound
            If Trim(Me.mnuReportItem(i).Tag) <> "" Then
                If Split(Me.mnuReportItem(i).Tag, ",")(1) = "ZL1_INSIDE_1261_2" Or Split(Me.mnuReportItem(i).Tag, ",")(1) = "ZL1_INSIDE_1261_3" Then
                    lngCount = lngCount + 1
                    If lngCount = Me.mnuReportItem.Count Then
                        Me.mnuReport.Visible = False
                    Else
                        Me.mnuReportItem(i).Visible = False
                    End If
                End If
            End If
        Next
    End If
    
    If mstrPrivs Like "*西成药*" Then
        bln收入项目 = IIf(zlDatabase.GetPara("西成药收入项目", 100, 1023) = 0, True, False)
    End If
    If mstrPrivs Like "*中成药*" And bln收入项目 = False Then
        bln收入项目 = IIf(zlDatabase.GetPara("中成药收入项目", 100, 1023) = 0, True, False)
    End If
    If mstrPrivs Like "*中成药*" And bln收入项目 = False Then
        bln收入项目 = IIf(zlDatabase.GetPara("中草药收入项目", 100, 1023) = 0, True, False)
    End If
    If bln收入项目 = True Then
        '模块公共参数已经调整到药品参数设置模块，目前没有私有或本机参数，暂时屏蔽参数设置界面
        MsgBox "请到药品参数设置模块设置各材质对应的收入项目！", vbInformation, gstrSysName
'        frmMediPara.ShowMe mstrPrivs, Me
        If gblnIncomeItem = False Then
            Unload Me
        End If
    End If
    
    If mbln自管药 = True Then
        tabContent.TabVisible(1) = False
        tabContent.TabVisible(2) = False
        tabContent.TabVisible(3) = False
    End If
End Sub



Private Sub LogisticPlatformInterface()
'物流平台接口
    
    If gobjLogisticPlatform Is Nothing Then
        On Error Resume Next
        Set gobjLogisticPlatform = CreateObject("zlDrugPurchase.clsDrugPurchase")
        If err <> 0 Then
            mnuUploadDrugInfo.Visible = False
            err.Clear: On Error GoTo 0
            Exit Sub
        End If
        
    End If
    
    If mnuEditSptPacker.Visible = False Then mnuEditSptPacker.Visible = True
    mnuUploadDrugInfo.Visible = True
       
End Sub

Private Sub Form_Resize()
    Dim lngTools As Single, lngStatus As Single
    
    If WindowState = 1 Then Exit Sub
    lngTools = IIf(Me.clbThis.Visible, Me.clbThis.Height, 0)
    lngStatus = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    err = 0: On Error Resume Next
    
    With Me.picVBar
        .Top = lngTools
        .Height = Me.ScaleHeight - lngTools - lngStatus
        If .Left < 2000 Then .Left = 2000
        If .Left > Me.ScaleWidth - 4000 Then .Left = Me.ScaleWidth - 4000
    End With
    With Me.picHBar
        .Left = Me.picVBar.Left + Me.picVBar.Width
        .Width = Me.ScaleWidth - .Left
        If .Top < 2000 Then .Top = 2000
        If .Top > Me.ScaleHeight - lngStatus - 2500 Then .Top = Me.ScaleHeight - lngStatus - 2500
    End With
    With Me.picClass
        .Left = Me.ScaleLeft
        .Top = lngTools
        .Height = Me.ScaleHeight - picClass.Top - lngStatus
        .Width = Me.picVBar.Left - Me.picClass.Left
    End With
    
    With Me.lvwItems
        .Left = Me.picVBar.Left + Me.picVBar.Width
        .Top = lngTools
        .Height = Me.picHBar.Top - .Top
        .Width = Me.ScaleWidth - .Left
    End With
    
    With Me.tabContent
        .Left = Me.picVBar.Left + Me.picVBar.Width
        .Top = Me.picHBar.Top + Me.picHBar.Height
        .Height = Me.ScaleHeight - lngStatus - .Top + 15
        .Width = Me.ScaleWidth - .Left + 15
    End With
    
    With Me.fraComment(0)
        .Left = 90
        .Width = Me.tabContent.Width - .Left * 2
        .Top = Me.tabContent.Height - .Height - 50 '- 90
    End With
    With Me.fraComment(1)
        .Left = 90
        .Width = Me.tabContent.Width - .Left * 2
        .Top = Me.tabContent.Height - .Height - 60
    End With
    
    With Me.lvwSpecs
        .Left = 90
        .Top = 395
        .Width = Me.tabContent.Width - .Left * 2
        If lblComment(0).Caption = "" Then
            lvwSpecs.Height = tabContent.Height - lvwSpecs.Top - 50
        Else
            lvwSpecs.Height = tabContent.Height - lvwSpecs.Top - 50 - fraComment(0).Height
        End If
    End With
    With Me.hgdPrice
        .Left = 90
        .Top = 395
        .Width = Me.tabContent.Width - .Left * 2
        If lblComment(3).Caption = "" Then
            hgdPrice.Height = tabContent.Height - hgdPrice.Top - 50
        Else
            hgdPrice.Height = tabContent.Height - hgdPrice.Top - 350 - lblComment(3).Height
        End If
    End With
    
    With Me.hgdCost
        .Left = 90
        .Top = 395
        .Width = Me.tabContent.Width - .Left * 2
        .Height = Me.tabContent.Height - .Top - 50
    End With
    
    With Me.hgdCharge
        .Left = 90
        .Top = 395
        .Width = Me.tabContent.Width - .Left * 2
        .Height = Me.tabContent.Height - .Top - 50
    End With
    
    SetParent txtFind.hwnd, tlbThis.hwnd
    SetParent picFind.hwnd, tlbThis.hwnd
    txtFind.Left = Me.ScaleWidth - txtFind.Width - 200
    picFind.Left = txtFind.Left - 100 - picFind.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    Call SaveListViewState(Me.lvwItems, Me.Name & Val(Me.tvwClass.Tag), Me.lvwItems.View)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\分割", "横向", Me.picVBar.Left)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\分割", "纵向", Me.picHBar.Top)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示所有", IIf(mnuViewShowAll.Checked, 1, 0)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示停用目录", IIf(mnuViewList.Checked, 1, 0)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示停用项目", IIf(mnuViewStoped.Checked, 1, 0)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示历史价格", IIf(mnuViewPrices.Checked, 1, 0)
    mstrNodeClick西成药 = ""
    mstrNodeClick中成药 = ""
    mstrNodeClick中草药 = ""
    mstrItemClick西成药 = ""
    mstrItemClick中成药 = ""
    mstrItemClick中草药 = ""
    mstrNodeSelect西成药 = ""
    mstrNodeSelect中成药 = ""
    mstrNodeSelect中草药 = ""
    mstrKey = ""
    mstrFindValue = ""
    Set mrsFind = Nothing
End Sub

Private Sub hgdCharge_DblClick()
    Dim strCharge As String
    On Error GoTo ErrHandle
    
    If InStr(mstrPrivs, "费别设置") = 0 Then Exit Sub
    
    If Me.hgdCharge.Rows > 1 Then
        If Me.hgdCharge.TextMatrix(Me.hgdCharge.Rows - 1, 1) <> "" Then
            strCharge = Me.hgdCharge.TextMatrix(Me.hgdCharge.Row, 1)
        End If
    End If
    If mnuPriceChargeSet.Enabled = True Then
'        frmChargeSortItemEdit.ShowMe Me, 3, strCharge, Val(hgdCharge.RowData(hgdCharge.Row)), hgdCharge.TextMatrix(Me.hgdCharge.Row, 0)
        frmSetExpense.ShowMe Me, Val(hgdCharge.RowData(hgdCharge.Row)), hgdCharge.TextMatrix(Me.hgdCharge.Row, 0)
        
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub hgdCost_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    Call PopupMenu(Me.mnuPrice, 2)
End Sub


Private Sub hgdPrice_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    
    Call PopupMenu(Me.mnuPrice, 2)
End Sub

Private Sub hgdPrice_RowColChange()
    Dim bln招标 As Boolean
    Dim rsCheck As New ADODB.Recordset
    If Val(hgdPrice.RowData(hgdPrice.Row)) = 0 Then Exit Sub
    
    On Error GoTo ErrHandle
    gstrSql = "Select Nvl(招标药品,0) 招标药品 From 药品规格 Where 药品ID=(Select 收费细目ID From 收费价目 Where ID=[1]" & _
            GetPriceClassString("") & ")"
    
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSql, "判断当前药品是否为招标药品", Val(hgdPrice.RowData(hgdPrice.Row)))
    
    bln招标 = (rsCheck!招标药品 = 1)
    
    With Me.hgdPrice
        Me.lblComment(3).Caption = "1、" & _
            .TextMatrix(.Row, colPrice药价类型) & "药品，" & _
            IIf(bln招标, "中标价格", "指导批价") & .TextMatrix(.Row, colPrice指导批价) & "元/" & .TextMatrix(.Row, colPrice单位) & "，" & _
            "采购扣率" & .TextMatrix(.Row, colPrice扣率) & "%。"
        Me.lblComment(4).Caption = "2、" & _
            "指导售价" & .TextMatrix(.Row, colPrice指导售价) & "元/" & .TextMatrix(.Row, colPrice单位) & "，" & _
            "指导差率" & .TextMatrix(.Row, colPrice指导差率) & "%，" & _
            IIf(Val(.TextMatrix(.Row, colPrice屏蔽费别)) = 0, "根据病人身份费别进行优惠或加价。", "不受病人身份费别影响。")
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    If Val(Me.tvwClass.Tag) < 2 Then
        With frmMediItem
            .Tag = IIf(Me.tvwClass.Tag = 0, 1, 2)
            .cmdCancel.Tag = "查阅"
            .lng分类id = Mid(Me.tvwClass.SelectedItem.Key, 2)
            .lng药名id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            .Show 1, Me
        End With
    ElseIf Val(Me.tvwClass.Tag) = 2 Then
        With frmMediHerbalItem
            .Tag = 3
            .cmdCancel.Tag = "查阅"
            .lng分类id = Mid(Me.tvwClass.SelectedItem.Key, 2)
            .lng药名id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            .strPrivs = Me.mstrPrivs
            .Show 1, Me
        End With
    ElseIf Val(Me.tvwClass.Tag) = 4 Then
        Set objItem = Me.lvwItems.SelectedItem
        If objItem Is Nothing Then
            Exit Sub
        End If
        If mstrType <> "7" Then
             With frmMediItem
                .Tag = objItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1)
                .cmdCancel.Tag = "查阅"
                .lng分类id = objItem.SubItems(Me.lvwItems.ColumnHeaders("_分类id").Index - 1)
                .lng药名id = Mid(Me.lvwItems.SelectedItem.Key, 2)
                .strPrivs = Me.mstrPrivs
                .Show 1, Me
            End With
        Else
            With frmMediHerbalItem
                .Tag = objItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1)
                .cmdCancel.Tag = "查阅"
                .lng分类id = objItem.SubItems(Me.lvwItems.ColumnHeaders("_分类id").Index - 1)
                .lng药名id = Mid(Me.lvwItems.SelectedItem.Key, 2)
                .strPrivs = Me.mstrPrivs
                .Show 1, Me
            End With
        End If
    End If
End Sub

Private Sub lvwItems_GotFocus()
    Set objItem = Me.lvwItems.SelectedItem
    If objItem Is Nothing Then
        Exit Sub
    End If
        
    If lvwItems.ListItems.Count = 0 Then Exit Sub
    If Val(Me.tvwClass.Tag) = 0 And InStr(1, mstrPrivs, "管理西成药") = 0 Then Exit Sub
    If Val(Me.tvwClass.Tag) = 1 And InStr(1, mstrPrivs, "管理中成药") = 0 Then Exit Sub
    If Val(Me.tvwClass.Tag) = 2 And InStr(1, mstrPrivs, "管理中草药") = 0 Then Exit Sub
    If Val(Me.tvwClass.Tag) = 4 And InStr(1, mstrPrivs, "管理西成药") = 0 _
        And InStr(1, mstrPrivs, "管理中成药") = 0 _
        And InStr(1, mstrPrivs, "管理中草药") = 0 Then
        Exit Sub
    End If
    
    '设置药品卡片的启用、停用标志
    If lvwItems.SelectedItem.Icon = "成药S" Or lvwItems.SelectedItem.Icon = "草药S" Then
        If mnuEditStart.Visible = True Then mnuEditStart.Enabled = True
        mnuEditStop.Enabled = False
        tlbThis.Buttons("Start").Enabled = mnuEditStart.Enabled
        tlbThis.Buttons("Stop").Enabled = False
    Else
        mnuEditStart.Enabled = False
        If mnuEditStop.Visible = True Then mnuEditStop.Enabled = True
        tlbThis.Buttons("Start").Enabled = False
        tlbThis.Buttons("Stop").Enabled = mnuEditStop.Enabled
    End If
    
    If Val(Me.tvwClass.Tag) = 0 Then
        If InStr(1, mstrPrivs, "管理西成药品种") = 0 Then
            Me.mnuEditItemAdd.Enabled = False
            Me.mnuEditItemMod.Enabled = False
            Me.mnuEditItemDel.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            tlbThis.Buttons("Start").Enabled = False
            tlbThis.Buttons("Stop").Enabled = False
        End If
    ElseIf Val(Me.tvwClass.Tag) = 1 Then
        If InStr(1, mstrPrivs, "管理中成药品种") = 0 Then
            Me.mnuEditItemAdd.Enabled = False
            Me.mnuEditItemMod.Enabled = False
            Me.mnuEditItemDel.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            tlbThis.Buttons("Start").Enabled = False
            tlbThis.Buttons("Stop").Enabled = False
        End If
    ElseIf Val(Me.tvwClass.Tag) = 2 Then
        If InStr(1, mstrPrivs, "管理中草药品种") = 0 Then
            Me.mnuEditItemAdd.Enabled = False
            Me.mnuEditItemMod.Enabled = False
            Me.mnuEditItemDel.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            tlbThis.Buttons("Start").Enabled = False
            tlbThis.Buttons("Stop").Enabled = False
        End If
    ElseIf Val(Me.tvwClass.Tag) = 4 Then
        If objItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1) = "1" Then
            If InStr(1, mstrPrivs, "管理西成药品种") = 0 Then
                Me.mnuEditItemAdd.Enabled = False
                Me.mnuEditItemMod.Enabled = False
                Me.mnuEditItemDel.Enabled = False
                mnuEditStart.Enabled = False
                mnuEditStop.Enabled = False
                tlbThis.Buttons("Start").Enabled = False
                tlbThis.Buttons("Stop").Enabled = False
            End If
        ElseIf objItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1) = "2" Then
            If InStr(1, mstrPrivs, "管理中成药品种") = 0 Then
                Me.mnuEditItemAdd.Enabled = False
                Me.mnuEditItemMod.Enabled = False
                Me.mnuEditItemDel.Enabled = False
                mnuEditStart.Enabled = False
                mnuEditStop.Enabled = False
                tlbThis.Buttons("Start").Enabled = False
                tlbThis.Buttons("Stop").Enabled = False
            End If
        ElseIf objItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1) = "3" Then
            If InStr(1, mstrPrivs, "管理中草药品种") = 0 Then
                Me.mnuEditItemAdd.Enabled = False
                Me.mnuEditItemMod.Enabled = False
                Me.mnuEditItemDel.Enabled = False
                mnuEditStart.Enabled = False
                mnuEditStop.Enabled = False
                tlbThis.Buttons("Start").Enabled = False
                tlbThis.Buttons("Stop").Enabled = False
            End If
        End If
    End If
    
End Sub

Private Sub lvwItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim strSqlCon As String
    Dim strCaption As String
    Dim str售价记录 As String
    
    err = 0: On Error GoTo ErrHand
    strCaption = lblComment(0).Caption
    str售价记录 = lblComment(3).Caption
    If Item.Index <> 1 Then '第一条记录默认是选中的不用再次选中了
        If mstr分类 = "0" Then
            mstrItemClick西成药 = Item.Key
            mstrNodeClick西成药 = tvwClass.SelectedItem.Key
        ElseIf mstr分类 = "1" Then
            mstrItemClick中成药 = Item.Key
            mstrNodeClick中成药 = tvwClass.SelectedItem.Key
        ElseIf mstr分类 = "2" Then
            mstrItemClick中草药 = Item.Key
            mstrNodeClick中草药 = tvwClass.SelectedItem.Key
        End If
    End If
    
    '----------填写规格-----------------
   
    gstrSql = "select Distinct I.ID,I.编码,I.规格,I.产地 as 厂牌,S.原产地,N.名称 as 商品名,I.费用类型 as 医保类型,S.药品来源," & _
            "        decode(I.服务对象,1,'门诊',2,'住院',3,'门诊和住院','不直接应用于病人') as 服务对象," & _
            "        decode(S.自制药品,1,'√',' ') as 自制,decode(S.协定药品,1,'√',' ') 协定," & _
            "        decode(S.招标药品,1,'√',' ') 招标,decode(S.中药形态,1,'中药饮片',2,'免煎剂','散装') 中药形态," & _
            "        S.批准文号,Nvl(I.是否变价,0) 是否变价,Nvl(S.招标药品,0) 招标药品," & _
            "        nvl(I.撤档时间,to_date('3000-01-01','YYYY-MM-DD')) as 撤档时间,G.名称 合同单位,I.说明,I.备选码,I.站点 " & _
            " from 收费项目目录 I,药品规格 S,收费项目别名 N,(Select Id,名称 From 供应商 Where 末级 = 1 And substr(类型,1,1) = '1' And " & _
            " (撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) ) G " & _
            " where I.ID=S.药品ID and G.id(+)=S.合同单位id and I.ID=N.收费细目ID(+) And N.性质(+) = 3 and S.药名ID=[1] "
'            " where (I.站点 = '" & gstrNodeNo & "' Or I.站点 is Null) And I.ID=S.药品ID and G.id(+)=S.合同单位id and I.ID=N.收费细目ID(+) And N.性质(+) = 3 and S.药名ID=[1] "
    If Me.mnuViewStoped.Checked = False Then
        gstrSql = gstrSql & " and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
    
    With rsTemp
        Me.lvwSpecs.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwSpecs.ListItems.Add(, "_" & !ID, IIf(IsNull(!规格), "", !规格))
            
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("厂牌").Index - 1) = IIf(IsNull(!厂牌), "", !厂牌)
            If Val(Me.tvwClass.Tag) = 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType = "7" Then
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("原产地").Index - 1) = IIf(IsNull(!原产地), "", !原产地)
            End If
            If Not (Val(Me.tvwClass.Tag) = 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType = "7") Then
                objItem.SubItems(Me.lvwSpecs.ColumnHeaders("商品名").Index - 1) = IIf(IsNull(!商品名), "", !商品名)
            End If
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("医保类型").Index - 1) = IIf(IsNull(!医保类型), "", !医保类型)
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("药品来源").Index - 1) = IIf(IsNull(!药品来源), "", !药品来源)
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("自制").Index - 1) = IIf(IsNull(!自制), "", !自制)
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("协定").Index - 1) = IIf(IsNull(!协定), "", !协定)
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("招标").Index - 1) = IIf(IsNull(!招标), "", !招标)
            If Val(Me.tvwClass.Tag) = 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType = "7" Then
                objItem.SubItems(Me.lvwSpecs.ColumnHeaders("中药形态").Index - 1) = IIf(IsNull(!中药形态), "", !中药形态)
            End If
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("批准文号").Index - 1) = IIf(IsNull(!批准文号), "", !批准文号)
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("合同单位").Index - 1) = IIf(IsNull(!合同单位), "", !合同单位)
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("说明").Index - 1) = IIf(IsNull(!说明), "", !说明)
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("备选码").Index - 1) = IIf(IsNull(!备选码), "", !备选码)
            objItem.SubItems(Me.lvwSpecs.ColumnHeaders("站点").Index - 1) = IIf(IsNull(!站点), "", !站点)
            'If Val(Me.tvwClass.Tag) < 2 Or (Val(Me.tvwClass.Tag) = 3 And mstrType <> "7") Then
                objItem.SubItems(Me.lvwSpecs.ColumnHeaders("服务对象").Index - 1) = IIf(IsNull(!服务对象), "", !服务对象)
            'End If
            
            If Format(!撤档时间, "YYYY-MM-DD") = "3000-01-01" Then
                If Val(Me.tvwClass.Tag) = 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType = "7" Then
                    objItem.Icon = "草规U": objItem.SmallIcon = "草规U"
                Else
                    objItem.Icon = "规格U": objItem.SmallIcon = "规格U"
                End If
            Else
                If Val(Me.tvwClass.Tag) = 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType = "7" Then
                    objItem.Icon = "草规S": objItem.SmallIcon = "草规S"
                Else
                    objItem.Icon = "规格S": objItem.SmallIcon = "规格S"
                End If
                
                objItem.ForeColor = mconColor_Stop
                For intCount = 1 To Me.lvwSpecs.ColumnHeaders.Count - 1
                    objItem.ListSubItems(intCount).ForeColor = mconColor_Stop
                Next
            End If

            '如果是招标药品，用颜色区分是否是时价还是定价药品
            If !招标药品 = 1 Then
                objItem.ListSubItems(1).ForeColor = IIf(!是否变价 = 0, &H800000, &H800080)
            Else
                objItem.ListSubItems(1).ForeColor = IIf(!是否变价 = 0, &H0, &H40&)
            End If
            .MoveNext
        Loop
    End With
    If Me.lvwSpecs.ListItems.Count > 0 Then
        If Me.lvwSpecs.SelectedItem Is Nothing Then Me.lvwSpecs.ListItems(1).Selected = True
        Call lvwSpecs_ItemClick(Me.lvwSpecs.SelectedItem)
        mnuEditSpecUnit.Enabled = True
    Else
        mnuEditSpecUnit.Enabled = False
        For intCount = Me.lblComment.LBound To Me.lblComment.UBound
            Me.lblComment(intCount).Caption = ""
        Next
    End If
    
    '----------填写售价-----------------
    With Me.hgdPrice
        .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
        For intCol = 0 To .Cols - 1
            .TextMatrix(.FixedRows, intCol) = ""
        Next
    End With
    
    gstrSql = "select P.ID,decode(I.是否变价,1,'时价','定价') as 药价类型,nvl(S.指导批发价,0) as 指导批价,nvl(S.扣率,0) as 扣率," & _
            "        nvl(S.指导零售价,0) as 指导售价,nvl(S.指导差价率,0) as 指导差率,nvl(I.屏蔽费别,0)  as 屏蔽费别," & _
            "        decode(sign(P.执行日期-sysdate),1,1,decode(sign(P.终止日期-sysdate),-1,-1,0)) as 执行情况," & _
            "        '['||I.编码||']'||I.名称||' '||I.规格||' '||I.产地 as 药品,I.计算单位 as 单位,S.药库单位,Nvl(S.药库包装,1) 药库包装," & _
            "        P.现价 as 售价,U.名称 as 收入项目,P.调价说明," & _
            "        to_char(P.执行日期,'YYYY-MM-DD HH24:MI:SS') as 执行日期,I.ID 药品ID,P.No 调价No " & _
            " from 收费价目 P,收入项目 U,收费项目目录 I,药品规格 S" & _
            " where P.收费细目ID=I.ID and P.收入项目ID=U.ID and I.ID=S.药品ID" & _
            "       And S.药名ID=[1] " & GetPriceClassString("P")
    If Me.mnuViewPrices.Checked = False Then
        gstrSql = gstrSql & " and (P.终止日期 is null or P.终止日期>=sysdate)"
    End If
    If Me.mnuViewStoped.Checked = False Then
        gstrSql = gstrSql & " and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    End If
    gstrSql = gstrSql & " order by I.编码,P.执行日期 desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
    
    With rsTemp
        Me.hgdPrice.Redraw = False
        If .BOF Or .EOF Then
            With Me.hgdPrice
                .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
                For intCol = 0 To .Cols - 1
                    .TextMatrix(.FixedRows, intCol) = ""
                Next
            End With
        Else
            Me.hgdPrice.Rows = Me.hgdPrice.FixedRows + .RecordCount
        End If
        Do While Not .EOF
            Me.hgdPrice.RowData(.AbsolutePosition) = !ID
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPrice药价类型) = !药价类型
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPrice指导批价) = Format(!指导批价 * IIf(mint药库单位 = 0, 1, !药库包装), mstrCostFormat)
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPrice扣率) = Format(!扣率, "0.00000;-0.00000;0")
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPrice指导售价) = Format(!指导售价 * IIf(mint药库单位 = 0, 1, !药库包装), mstrPriceFormat)
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPrice指导差率) = Format(!指导差率, "0.00000;-0.00000;0")
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPrice屏蔽费别) = !屏蔽费别
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPrice执行情况) = !执行情况
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPrice调价NO) = IIf(IsNull(!调价No), "", !调价No)
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPrice药品) = !药品
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPrice单位) = IIf(mint药库单位 = 0, !单位, !药库单位)
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPrice售价) = Format(!售价 * IIf(mint药库单位 = 0, 1, !药库包装), mstrPriceFormat)
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPrice收入项目) = !收入项目
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPrice说明) = IIf(IsNull(!调价说明), "", !调价说明)
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPrice执行日期) = !执行日期
            Me.hgdPrice.TextMatrix(.AbsolutePosition, colPrice药品ID) = !药品id
            Me.hgdPrice.Row = .AbsolutePosition
            For intCol = 0 To Me.hgdPrice.Cols - 1
                Me.hgdPrice.Col = intCol
                Select Case !执行情况
                Case -1
                    Me.hgdPrice.CellBackColor = RGB(240, 240, 240)
                Case 0
                    Me.hgdPrice.CellBackColor = RGB(255, 255, 255)
                Case 1
                    Me.hgdPrice.CellBackColor = RGB(225, 255, 255)
                End Select
            Next
            .MoveNext
        Loop
        Me.hgdPrice.Row = Me.hgdPrice.FixedRows
        If Val(Me.tvwClass.Tag) < 2 Or (Val(Me.tvwClass.Tag) = 4 And mstrType <> "7") Then
            If Me.hgdPrice.ColWidth(colPrice药品) = 0 Or Me.hgdPrice.ColWidth(colPrice单位) = 0 Then
                Me.hgdPrice.ColWidth(colPrice药品) = 3500
                Me.hgdPrice.ColWidth(colPrice单位) = 550
            End If
        ElseIf Val(Me.tvwClass.Tag) = 2 Or (Val(Me.tvwClass.Tag) = 4 And mstrType = "7") Then
            Me.hgdPrice.ColWidth(colPrice药品) = 0
            Me.hgdPrice.ColWidth(colPrice单位) = 0
        End If
        Me.hgdPrice.Redraw = True
    End With
    
    Call hgdPrice_RowColChange
    
    '取成本价调价记录
    mlngCurrDrug = Val(Mid(Item.Key, 2))
    If Me.tabContent.Tab = 2 Then
        Call GetCostAdjust(mlngCurrDrug)
    End If
    
    '取费别等级设置
    Call GetChargeSet(mlngCurrDrug)
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Call lvwItems_GotFocus
    If lvwSpecs.SelectedItem Is Nothing Then
        mnuEditSpecMod.Enabled = False
        mnuEditSpecDel.Enabled = False
    ElseIf lvwSpecs.SelectedItem.ForeColor = vbRed Then
        mnuEditSpecMod.Enabled = False
    Else
        mnuEditSpecMod.Enabled = True
        mnuEditSpecDel.Enabled = True
    End If
    If lvwItems.SelectedItem.ForeColor = vbRed Then
        mnuEditSpecAdd.Enabled = False
        mnuEditItemMod.Enabled = False
    Else
        mnuEditSpecAdd.Enabled = True
    End If
    
    If lblComment(0).Caption = "" Then
        lvwSpecs.Height = tabContent.Height - lvwSpecs.Top - 50
    Else
        lvwSpecs.Height = tabContent.Height - lvwSpecs.Top - 50 - fraComment(0).Height
    End If

    If lblComment(3).Caption = "" Then
        hgdPrice.Height = tabContent.Height - hgdPrice.Top - 50
    Else
        hgdPrice.Height = tabContent.Height - hgdPrice.Top - 350 - lblComment(3).Height
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_DblClick
End Sub

Private Sub lvwItems_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    Call zlPopupEditMenu(1, True)
End Sub

Private Sub lvwSpecs_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwSpecs.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwSpecs.SortOrder = IIf(Me.lvwSpecs.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwSpecs.SortKey = ColumnHeader.Index - 1
        Me.lvwSpecs.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwSpecs_DblClick()
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    If Me.lvwSpecs.SelectedItem Is Nothing Then Exit Sub
    If Val(Me.tvwClass.Tag) = 2 Or mstrType = "7" Then
        With frmMediHerbalSpec
            .stbSpec.Tag = "查阅"
            .lng药名id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            .lng药品ID = Mid(Me.lvwSpecs.SelectedItem.Key, 2)
            .strPrivs = Me.mstrPrivs
            .Show 1, Me
        End With
    Else
        With frmMediSpec
            .stbSpec.Tag = "查阅"
            .lng药名id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            .lng药品ID = Mid(Me.lvwSpecs.SelectedItem.Key, 2)
            .strPrivs = Me.mstrPrivs
            .Show 1, Me
        End With
    End If
End Sub

Private Sub lvwSpecs_GotFocus()
    Set objItem = Me.lvwItems.SelectedItem
    If objItem Is Nothing Then
        Exit Sub
    End If
    
    If lvwSpecs.ListItems.Count = 0 Then Exit Sub
    If Val(Me.tvwClass.Tag) = 0 And InStr(1, mstrPrivs, "管理西成药") = 0 Then Exit Sub
    If Val(Me.tvwClass.Tag) = 1 And InStr(1, mstrPrivs, "管理中成药") = 0 Then Exit Sub
    If Val(Me.tvwClass.Tag) = 2 And InStr(1, mstrPrivs, "管理中草药") = 0 Then Exit Sub
    If Val(Me.tvwClass.Tag) = 4 And InStr(1, mstrPrivs, "管理西成药") = 0 _
        And InStr(1, mstrPrivs, "管理中成药") = 0 _
        And InStr(1, mstrPrivs, "管理中草药") = 0 Then
        Exit Sub
    End If
    
    '设置药品卡片的启用、停用标志
    If lvwSpecs.SelectedItem.Icon = "规格S" Or lvwSpecs.SelectedItem.Icon = "草规S" Then
        If mnuEditStart.Visible = True Then mnuEditStart.Enabled = True
        mnuEditStop.Enabled = False
        tlbThis.Buttons("Start").Enabled = mnuEditStart.Enabled
        tlbThis.Buttons("Stop").Enabled = False
    Else
        mnuEditStart.Enabled = False
        If mnuEditStop.Visible = True Then mnuEditStop.Enabled = True
        tlbThis.Buttons("Start").Enabled = False
        tlbThis.Buttons("Stop").Enabled = mnuEditStop.Enabled
    End If
    
    If Val(Me.tvwClass.Tag) = 0 Then
        If InStr(1, mstrPrivs, "管理西成药规格") = 0 Then
            Me.mnuEditSpecAdd.Enabled = False
            Me.mnuEditSpecMod.Enabled = False
            Me.mnuEditSpecDel.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            tlbThis.Buttons("Start").Enabled = False
            tlbThis.Buttons("Stop").Enabled = False
        End If
    ElseIf Val(Me.tvwClass.Tag) = 1 Then
        If InStr(1, mstrPrivs, "管理中成药规格") = 0 Then
            Me.mnuEditSpecAdd.Enabled = False
            Me.mnuEditSpecMod.Enabled = False
            Me.mnuEditSpecDel.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            tlbThis.Buttons("Start").Enabled = False
            tlbThis.Buttons("Stop").Enabled = False
        End If
    ElseIf Val(Me.tvwClass.Tag) = 2 Then
        If InStr(1, mstrPrivs, "管理中草药规格") = 0 Then
            Me.mnuEditSpecAdd.Enabled = False
            Me.mnuEditSpecMod.Enabled = False
            Me.mnuEditSpecDel.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            tlbThis.Buttons("Start").Enabled = False
            tlbThis.Buttons("Stop").Enabled = False
        End If
    ElseIf Val(Me.tvwClass.Tag) = 4 Then
        If objItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1) = "1" Then
            If InStr(1, mstrPrivs, "管理西成药规格") = 0 Then
                Me.mnuEditSpecAdd.Enabled = False
                Me.mnuEditSpecMod.Enabled = False
                Me.mnuEditSpecDel.Enabled = False
                mnuEditStart.Enabled = False
                mnuEditStop.Enabled = False
                tlbThis.Buttons("Start").Enabled = False
                tlbThis.Buttons("Stop").Enabled = False
            End If
        ElseIf objItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1) = "2" Then
            If InStr(1, mstrPrivs, "管理中成药规格") = 0 Then
                Me.mnuEditSpecAdd.Enabled = False
                Me.mnuEditSpecMod.Enabled = False
                Me.mnuEditSpecDel.Enabled = False
                mnuEditStart.Enabled = False
                mnuEditStop.Enabled = False
                tlbThis.Buttons("Start").Enabled = False
                tlbThis.Buttons("Stop").Enabled = False
            End If
        ElseIf objItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1) = "3" Then
            If InStr(1, mstrPrivs, "管理中草药规格") = 0 Then
                Me.mnuEditSpecAdd.Enabled = False
                Me.mnuEditSpecMod.Enabled = False
                Me.mnuEditSpecDel.Enabled = False
                mnuEditStart.Enabled = False
                mnuEditStop.Enabled = False
                tlbThis.Buttons("Start").Enabled = False
                tlbThis.Buttons("Stop").Enabled = False
            End If
        End If
    End If
End Sub

Private Sub lvwSpecs_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim rsData As ADODB.Recordset
    
    err = 0: On Error GoTo ErrHand
    
    '如果已到执行日期而价格未执行，执行计算过程
'    gstrSql = " Select ID From 收费价目 Where 收费细目ID=[1] And 变动原因=0" & GetPriceClassString("")
'
'    Set rsData = zlDatabase.OpenSQLRecord(gstrSql, "检查未执行的价格", Val(Mid(Item.Key, 2)))
'
'    With rsData
'        If Not .EOF Then
'            If Not IsNull(!ID) Then
'                gstrSql = "zl_药品收发记录_Adjust(" & Val(!ID) & ")"
'                Call zlDatabase.ExecuteProcedure(gstrSql, "产生药品价格调整记录")
'            End If
'        End If
'    End With
    
    gstrSql = "zl_药品收发记录_Adjust(" & Val(Mid(Item.Key, 2)) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSql, "产生药品价格调整记录")
        
    gstrSql = "select I.计算单位||decode(I.计算单位,O.计算单位,'','(='||decode(sign(S.剂量系数-1),-1,'0','')||to_char(S.剂量系数)||O.计算单位||')') as 售价单位," & _
            "        S.门诊单位||decode(S.门诊单位,I.计算单位,'','(='||decode(sign(S.门诊包装-1),-1,'0','')||to_char(S.门诊包装)||I.计算单位||')') as 门诊单位," & _
            "        S.住院单位||decode(S.住院单位,I.计算单位,'','(='||decode(sign(S.住院包装-1),-1,'0','')||to_char(S.住院包装)||I.计算单位||')') as 住院单位," & _
            "        S.药库单位||decode(S.药库单位,I.计算单位,'','(='||decode(sign(S.药库包装-1),-1,'0','')||to_char(S.药库包装)||I.计算单位||')') as 药库单位," & _
            "        nvl(S.药库分批,0) as 药库分批,nvl(S.药房分批,0) as 药房分批,nvl(S.最大效期,0) as 最大效期,nvl(S.住院可否分零,0) as 可否分零,Nvl(To_Char(I.撤档时间,'yyyy-MM-dd'),'3000-01-01') As 撤档时间" & _
            " from 药品规格 S,收费项目目录 I,诊疗项目目录 O" & _
            " where S.药品ID=I.ID and S.药名ID=O.ID" & _
            "       and S.药品id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
        
    With rsTemp
        If Val(Me.tvwClass.Tag) = 2 Or mstrType = "7" Then
            Me.lblComment(0).Caption = "1、售价单位：" & !售价单位 & "； 药房单位：" & !门诊单位 & "； 药库单位：" & !药库单位 & "。"
        Else
            Me.lblComment(0).Caption = "1、售价单位：" & !售价单位 & "； 门诊单位：" & !门诊单位 & "； 住院单位：" & !住院单位 & "； 药库单位：" & !药库单位 & "。"
        End If
        
        If !药库分批 = 0 Then
            Me.lblComment(1).Caption = "2、该药品不进行分批管理。"
        Else
            If !药房分批 = 0 Then
                Me.lblComment(1).Caption = "2、该药品在药库中分批管理。"
            Else
                Me.lblComment(1).Caption = "2、该药品在药库药房都需要分批管理。"
            End If
            If !最大效期 = 0 Then
                Me.lblComment(1).Caption = Me.lblComment(1).Caption & "但不进行效期跟踪。"
            Else
                Me.lblComment(1).Caption = Me.lblComment(1).Caption & "最长保持期" & !最大效期 & "月。"
            End If
        End If
        Select Case !可否分零
        Case 0
            Me.lblComment(2).Caption = "3、该药品允许分零应用。"
        Case 1
            Me.lblComment(2).Caption = "3、该药品不允许分零应用。"
        Case 2
            Me.lblComment(2).Caption = "3、该药品为一次性药品。"
        Case Is < 0
            Me.lblComment(2).Caption = "3、该药品分零后" & Abs(!可否分零) & "天内使用有效。"
        Case Else
        End Select
    End With
    Call lvwSpecs_GotFocus
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lvwSpecs_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Me.lvwSpecs.SelectedItem Is Nothing Then Exit Sub
    Call lvwSpecs_DblClick
End Sub

Private Sub lvwSpecs_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    
    Call zlPopupEditMenu(2, True)
End Sub

Private Sub mnuClassAdd_Click()
    Dim intTab As Integer
    
    intTab = tabContent.Tab
    If Val(Me.tvwClass.Tag) = 4 Then            '=3是显示过滤结果的状态，不允许编辑类别
        Exit Sub
    End If
    With frmClinicClass
        .lblKind.Tag = Val(Me.tvwClass.Tag) + 1
        If Me.tvwClass.SelectedItem Is Nothing Then
            .txtParent.Tag = 0
        Else
            .txtParent.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
            If tvwClass.SelectedItem.Text <> "" Then
                .txtParent.Text = tvwClass.SelectedItem.Text
            End If
        End If
        .Tag = "增加"
        .Show 1, Me
    End With
    If gblnCancel = False Then
        If Me.tvwClass.SelectedItem Is Nothing Then
            Call zlRefClasses
        Else
            Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Key, 2))
        End If
    End If
    tabContent.Tab = intTab
End Sub

Private Sub mnuClassDel_Click()
    If Val(Me.tvwClass.Tag) = 4 Then           '=3是显示过滤结果的状态，不允许编辑类别
        Exit Sub
    End If
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("真的删除该分类“" & Me.tvwClass.SelectedItem.Text & "”吗", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    err = 0: On Error GoTo ErrHand
    gstrSql = "zl_诊疗分类目录_delete(" & Mid(Me.tvwClass.SelectedItem.Key, 2) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    
    Dim strParentKey As String
    If Me.tvwClass.SelectedItem.Next Is Nothing Then
        If Me.tvwClass.SelectedItem.Parent Is Nothing Then
            Call zlRefClasses
        Else
            strParentKey = Me.tvwClass.SelectedItem.Parent.Key
            Call Me.tvwClass.Nodes.Remove(Me.tvwClass.SelectedItem.Key)
            If Me.tvwClass.Nodes(strParentKey).Children = 0 Then
                Call zlRefClasses(Mid(Me.tvwClass.Nodes(strParentKey).Key, 2))
            Else
                Call zlRefClasses(Mid(Me.tvwClass.Nodes(strParentKey).Child.Key, 2))
            End If
        End If
    Else
        Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Next.Key, 2))
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuClassMod_Click()
    Dim intTab As Integer   '记录以前所选择的页面
    
    intTab = tabContent.Tab
    If Val(Me.tvwClass.Tag) = 4 Then            '=3是显示过滤结果的状态，不允许编辑类别
        Exit Sub
    End If
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    With frmClinicClass
        .lblKind.Tag = Val(Me.tvwClass.Tag) + 1
        If Me.tvwClass.SelectedItem.Parent Is Nothing Then
            .txtParent.Tag = 0
            .txtParent.Text = "(无)"
            .txtUpCode.Text = ""
            .txtCode.Text = Mid(Split(Me.tvwClass.SelectedItem.Text, "]")(0), 2)
            .txtCode.MaxLength = Len(.txtCode.Text)
            .txtCode.Tag = .txtCode.MaxLength
        Else
            .txtParent.Tag = Mid(Me.tvwClass.SelectedItem.Parent.Key, 2)
            .txtParent.Text = Me.tvwClass.SelectedItem.Parent.Text
            .txtUpCode.Text = Mid(Split(Me.tvwClass.SelectedItem.Parent.Text, "]")(0), 2)
            .txtCode.Text = Mid(Split(Me.tvwClass.SelectedItem.Text, "]")(0), Len(.txtUpCode.Text) + 2)
            .txtCode.MaxLength = Len(.txtCode.Text)
            .txtCode.Tag = .txtCode.MaxLength
        End If
        .txtName = Split(Me.tvwClass.SelectedItem.Text, "]")(1)
        .txtSymbol = Me.tvwClass.SelectedItem.Tag
        .Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
        .Show 1, Me
    End With
    If gblnCancel = False Then
        Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Key, 2))
    End If
    tabContent.Tab = intTab
End Sub

Private Sub mnuClassStar_Click()
    '停用分类、子分类、分类下品种及规格
    
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    
    frmMediClassReuse.ShowForm Val(Mid(tvwClass.SelectedItem.Key, 2)), Me.tvwClass.Tag
    
    Call zlRefClasses
End Sub

Private Sub mnuClassStop_Click()
    '停用分类、子分类、分类下品种及规格
    
    On Error GoTo ErrHand
    
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    
    If MsgBox("是否停用该分类及该分类下所有药品吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSql = "Zl_诊疗分类目录_药品分类停用(" & Val(Mid(tvwClass.SelectedItem.Key, 2)) & "," & Val(Me.tvwClass.Tag) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    
    Call zlRefClasses
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub mnuEditContrast_Click()
    frmMediContrast.ShowMe Me
End Sub

Private Sub mnuEditExcel_Click()
'    frmItemImport.ShowMe 2, Me
    frmImportFile.ShowMe 2, Me
'    Call zlRefClasses
    Call zlRefRecords
    If Me.tvwClass.SelectedItem Is Nothing Then
        Call zlRefClasses
    Else
        Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Key, 2))
    End If
End Sub

Private Sub mnuEditManFac_Click()
    Dim strType As String
    Dim str类型 As String
    Dim lng药品ID As String
    
    On Error Resume Next
    '厂家和批准文号设置
    With frmSetManfac
        
        If Me.tvwClass.Tag = 4 Or Me.tvwClass.Tag = 3 Then '单独处理过滤结果
            If Me.lvwItems.SelectedItem Is Nothing Then     '如果没有记录就退出，因为无法判断药品材质
                Exit Sub
            End If
            strType = Me.lvwItems.SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1)
            strType = Switch(strType = "1", "5", strType = "2", "6", strType = "3", "7")
            str类型 = strType
        Else
            str类型 = Switch(Me.tvwClass.Tag = "0", "5", Me.tvwClass.Tag = "1", "6", Me.tvwClass.Tag = "2", "7")
        End If
        If Me.lvwSpecs.SelectedItem Is Nothing Then
            lng药品ID = 0
        Else
            lng药品ID = Mid(Me.lvwSpecs.SelectedItem.Key, 2)
        End If
        .ShowMe str类型, Me.mstrPrivs, lng药品ID
    End With
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuEditRate_Click()
    frm加成率设置.ShowMe Me
End Sub

Private Sub mnuEditSendType_Click()
    frmMediSendType.Show vbModal, Me
End Sub

Private Sub mnuEditSpecBatch_Click()
    '规格修改
    frmBatchUpdate.ShowMe 2, mstrPrivs, mbln自管药
End Sub

Private Sub mnuEditSpecExp_Click()
    frmMediSpecExp.Show
End Sub

Private Sub mnuEditVariBatch_Click()
    '1是品种修改
    frmBatchUpdate.ShowMe 1, mstrPrivs, mbln自管药
End Sub

Private Sub mnuPriceChargeSet_Click()
    If Me.lvwSpecs.SelectedItem Is Nothing Then Exit Sub
'    frmChargeSortItemEdit.ShowMe Me, 3, "", Val(Mid(Me.lvwSpecs.SelectedItem.Key, 2)), Me.lvwSpecs.SelectedItem.Text
    
    frmSetExpense.ShowMe Me, Val(Mid(Me.lvwSpecs.SelectedItem.Key, 2)), Me.lvwSpecs.SelectedItem.Text
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuEditItemAdd_Click()
    Dim lng分类id As Long
    Dim lng药名id As Long
    Dim int类型 As Integer
    
    If Me.tvwClass.SelectedItem Is Nothing Then MsgBox "尚未设置分类,不能增删品种！", vbExclamation, gstrSysName: Exit Sub
    If Val(Me.tvwClass.Tag) < 2 Then
        With frmMediItem
            .Tag = IIf(Me.tvwClass.Tag = 0, 1, 2)
            .cmdCancel.Tag = "增加"
            .lng分类id = Mid(Me.tvwClass.SelectedItem.Key, 2)
            If Me.lvwItems.SelectedItem Is Nothing Then
                .lng药名id = 0
            Else
                .lng药名id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            End If
            .strPrivs = Me.mstrPrivs
            .lng抗生素 = 0
            .ShowMe mbln自管药, Me
        End With
    ElseIf Val(Me.tvwClass.Tag) = 2 Then
        '草药品种
        With frmMediHerbalItem
            .Tag = 3
            .cmdCancel.Tag = "增加"
            .lng分类id = Mid(Me.tvwClass.SelectedItem.Key, 2)
            If Me.lvwItems.SelectedItem Is Nothing Then
                .lng药名id = 0
            Else
                .lng药名id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            End If
            .strPrivs = Me.mstrPrivs
            .ShowMe mbln自管药, Me
        End With
    ElseIf Val(Me.tvwClass.Tag) = 3 Then
        If Not lvwItems.SelectedItem Is Nothing Then
            lng分类id = objItem.SubItems(Me.lvwItems.ColumnHeaders("_分类id").Index - 1)
            lng药名id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            int类型 = objItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1)
        Else
            Exit Sub
        End If
        If mstrType <> "7" Then
             With frmMediItem
                .lng抗生素 = 0
                .chk原研药.Value = 0
                .chk专利药.Value = 0
                .chk单独定价.Value = 0
                .Tag = int类型
                .cmdCancel.Tag = "增加"
                .lng分类id = lng分类id
                .lng药名id = lng药名id
                If Val(Me.tvwClass.Tag) = 3 Then
                    If Not Me.tvwClass.SelectedItem.Parent Is Nothing Then
                        If IsNumeric(Mid(Me.tvwClass.SelectedItem.Key, 2)) Then
                            If Me.tvwClass.SelectedItem.Parent.Key Like "_L*" Then
                                .lng抗生素 = Mid(Me.tvwClass.SelectedItem.Key, 2, 1)
                            ElseIf Me.tvwClass.SelectedItem.Parent.Key Like "_原研药" Then
                                .chk原研药.Value = 1
                            ElseIf Me.tvwClass.SelectedItem.Parent.Key Like "_专利药" Then
                                .chk专利药.Value = 1
                            ElseIf Me.tvwClass.SelectedItem.Parent.Key Like "_单独定价" Then
                                .chk单独定价.Value = 1
                            End If
                        Else
                            If Me.tvwClass.SelectedItem.Parent.Key Like "_抗菌药" Then
                                .lng抗生素 = Mid(Me.tvwClass.SelectedItem.Key, 7, 1)
                            End If
                        End If
                    Else
                        If Me.tvwClass.SelectedItem.Key Like "_抗菌药" Then
                            .lng抗生素 = 1
                        ElseIf Me.tvwClass.SelectedItem.Key Like "_原研药" Then
                            .chk原研药.Value = 1
                        ElseIf Me.tvwClass.SelectedItem.Key Like "_专利药" Then
                            .chk专利药.Value = 1
                        ElseIf Me.tvwClass.SelectedItem.Key Like "_单独定价" Then
                            .chk单独定价.Value = 1
                        End If
                    End If
                End If
                
                .strPrivs = Me.mstrPrivs
                .ShowMe mbln自管药, Me
            End With
        Else
            With frmMediHerbalItem
                .Tag = 3
                .cmdCancel.Tag = "增加"
                .lng分类id = lng分类id
                .lng药名id = lng药名id
                .strPrivs = Me.mstrPrivs
                .ShowMe mbln自管药, Me
            End With
        End If
    ElseIf Val(Me.tvwClass.Tag) = 4 Then        '单独处理过滤结果
        If (tvwClass.SelectedItem.Key Like "_L*" Or tvwClass.SelectedItem.Key Like "_A*") And lvwItems.SelectedItem Is Nothing Then
            Exit Sub
        ElseIf (tvwClass.SelectedItem.Key Like "_L*" Or tvwClass.SelectedItem.Key Like "_A*") And Not lvwItems.SelectedItem Is Nothing Then
            lng分类id = objItem.SubItems(Me.lvwItems.ColumnHeaders("_分类id").Index - 1)
            lng药名id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            int类型 = objItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1)
        ElseIf (Not tvwClass.SelectedItem.Key Like "_L*" Or tvwClass.SelectedItem.Key Like "_A*") And lvwItems.SelectedItem Is Nothing Then
            lng分类id = Mid(Me.tvwClass.SelectedItem.Key, IIf(Val(Me.tvwClass.Tag) = 3, 3, 2))
            lng药名id = 0
            int类型 = 1
        ElseIf (Not tvwClass.SelectedItem.Key Like "_L*" Or tvwClass.SelectedItem.Key Like "_A*") And Not lvwItems.SelectedItem Is Nothing Then
            lng分类id = Mid(Me.tvwClass.SelectedItem.Key, IIf(Val(Me.tvwClass.Tag) = 3, 3, 2))
            lng药名id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            int类型 = objItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1)
        End If
        If mstrType <> "7" Then
             With frmMediItem
                .Tag = int类型
                .cmdCancel.Tag = "增加"
                .lng分类id = lng分类id
                .lng药名id = lng药名id
                .strPrivs = Me.mstrPrivs
                .ShowMe mbln自管药, Me
            End With
        Else
            With frmMediHerbalItem
                .Tag = 3
                .cmdCancel.Tag = "增加"
                .lng分类id = lng分类id
                .lng药名id = lng药名id
                .strPrivs = Me.mstrPrivs
                .ShowMe mbln自管药, Me
            End With
        End If
    End If
    If gblnCancel = False Then
        Call zlRefRecords
    End If
End Sub

Private Sub mnuEditItemBill_Click()
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call frmClinicBill.ShowMe(Me, Mid(Me.lvwItems.SelectedItem.Key, 2))
    Call zlRefRecords(Mid(Me.lvwItems.SelectedItem.Key, 2))
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuEditItemDel_Click()
    Dim lngItem As Long
    Dim intCol As Integer
    Dim blnTrans As Boolean
    Dim rsSpec As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    With Me.lvwItems
        If .SelectedItem Is Nothing Then Exit Sub
        If MsgBox("真的删除“" & .SelectedItem.Text & "”吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        lngItem = Mid(.SelectedItem.Key, 2)
        gstrSql = "Select 药品ID From 药品规格 Where 药名ID=[1]"
        Set rsSpec = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItem)
        
        gcnOracle.BeginTrans
        blnTrans = True
        '删除诊疗项目目录
'        If Val(Me.tvwClass.Tag) < 2 Or (Val(Me.tvwClass.Tag) = 3 And mstrType <> "7") Then
            gstrSql = "zl_成药品种_DELETE(" & lngItem & ")"
'        ElseIf Val(Me.tvwClass.Tag) = 2 Or (Val(Me.tvwClass.Tag) = 3 And mstrType = "7") Then
'            gstrSql = "zl_草药药品_DELETE(" & lngItem & ")"
'        End If
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        
        '删除对应的收费项目目录
        Do While Not rsSpec.EOF
            gstrSql = "zl_成药规格_DELETE(" & rsSpec!药品id & ")"
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            rsSpec.MoveNext
        Loop
        gcnOracle.CommitTrans
        blnTrans = False
        
        '同步删除物流平台药品信息
        If Not gobjLogisticPlatform Is Nothing And rsSpec.RecordCount > 0 Then
            rsSpec.MoveFirst
            Do While Not rsSpec.EOF
                gobjLogisticPlatform.ClearDrugInfo rsSpec!药品id, 0
                rsSpec.MoveNext
            Loop
        End If
        
        Call .ListItems.Remove(.SelectedItem.Key)
        If .SelectedItem Is Nothing Then
            lvwSpecs.ListItems.Clear
            With Me.hgdPrice
                .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
                For intCol = 0 To .Cols - 1
                    .TextMatrix(.FixedRows, intCol) = ""
                Next
            End With
        Else
            Call lvwItems_ItemClick(.SelectedItem)
        End If
        
        '处理过滤结果：在返回药名ID串中去除已经删除的药名id
        Dim i As Integer
        Dim strAryDrugId() As String
        Dim strTmp As String
        
        If Val(Me.tvwClass.Tag) = 4 Then
            mstrDrugId = mstrDrugId & ","
            strAryDrugId = Split(mstrDrugId, ",")
            For i = 0 To UBound(strAryDrugId) - 1
                If strAryDrugId(i) <> CStr(lngItem) Then
                    strTmp = strTmp & strAryDrugId(i) & ","
                End If
            Next
            If Len(strTmp) > 1 Then
                strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
            Else
                strTmp = ""
            End If
            mstrDrugId = strTmp
        End If
        
    End With
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditItemMod_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then MsgBox "尚未设置分类,不能增删品种！", vbExclamation, gstrSysName: Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    If Me.lvwItems.SelectedItem.Icon = "成药S" Then
        MsgBox "不能对停用药品药品进行修改！", vbExclamation, gstrSysName
        Exit Sub
    End If
    If Val(Me.tvwClass.Tag) < 2 Then
        With frmMediItem
            .Tag = IIf(Me.tvwClass.Tag = 0, 1, 2)
            .cmdCancel.Tag = "修改"
            .lng分类id = Mid(Me.tvwClass.SelectedItem.Key, 2)
            .lng药名id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            .strPrivs = Me.mstrPrivs
            .lng抗生素 = 0
            .Show 1, Me
        End With
    ElseIf Val(Me.tvwClass.Tag) = 2 Then
        With frmMediHerbalItem
            .Tag = 3
            .cmdCancel.Tag = "修改"
            .lng分类id = Mid(Me.tvwClass.SelectedItem.Key, 2)
            .lng药名id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            .strPrivs = Me.mstrPrivs
            .Show 1, Me
        End With
    ElseIf Val(Me.tvwClass.Tag) = 3 Or Val(Me.tvwClass.Tag) = 4 Then        '单独处理过滤结果
        Set objItem = Me.lvwItems.SelectedItem
        If objItem Is Nothing Then
            Exit Sub
        End If
        If mstrType <> "7" Then
             With frmMediItem
                .Tag = objItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1)
                .cmdCancel.Tag = "修改"
                .lng分类id = objItem.SubItems(Me.lvwItems.ColumnHeaders("_分类id").Index - 1)
                .lng药名id = Mid(Me.lvwItems.SelectedItem.Key, 2)
                .strPrivs = Me.mstrPrivs
                If Val(Me.tvwClass.Tag) = 3 Then
                    If Not Me.tvwClass.SelectedItem.Parent Is Nothing Then
                        If IsNumeric(Mid(Me.tvwClass.SelectedItem.Key, 2)) Then
                            If Me.tvwClass.SelectedItem.Parent.Key Like "_L*" Then
                                .lng抗生素 = Mid(Me.tvwClass.SelectedItem.Key, 2, 1)
                            End If
                        Else
                            If Me.tvwClass.SelectedItem.Parent.Key Like "_抗菌药" Then
                                .lng抗生素 = Mid(Me.tvwClass.SelectedItem.Key, 7, 1)
                            End If
                        End If
                    End If
                End If
                
                .Show 1, Me
            End With
        Else
            With frmMediHerbalItem
                .Tag = 3
                .cmdCancel.Tag = "修改"
                .lng分类id = objItem.SubItems(Me.lvwItems.ColumnHeaders("_分类id").Index - 1)
                .lng药名id = Mid(Me.lvwItems.SelectedItem.Key, 2)
                .strPrivs = Me.mstrPrivs
                .Show 1, Me
            End With
        End If
    End If
    If gblnCancel = False Then
        If Not (Me.lvwItems.SelectedItem Is Nothing) Then Call zlRefRecords(Mid(Me.lvwItems.SelectedItem.Key, 2))
    End If
End Sub

Private Sub mnuEditItemPart_Click()
    Dim int用途分类 As Integer, lng药品ID As Long, bln编辑 As Boolean
    Dim strStationNo As String
    With frmServiceSectOffice
        Dim strType As String
        If Me.tvwClass.Tag = 4 Or Me.tvwClass.Tag = 3 Then '单独处理过滤结果
            If Me.lvwItems.SelectedItem Is Nothing Then     '如果没有记录就退出，因为无法判断药品材质
                Exit Sub
            End If
            int用途分类 = CInt(Me.lvwItems.SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1))
            int用途分类 = Switch(int用途分类 = 1, 5, int用途分类 = 2, 6, int用途分类 = 3, 7)
        Else
            int用途分类 = Switch(Me.tvwClass.Tag = "0", 5, Me.tvwClass.Tag = "1", 6, Me.tvwClass.Tag = "2", 7)
        End If
        If Me.lvwSpecs.SelectedItem Is Nothing Then
            lng药品ID = 0
        Else
            lng药品ID = Mid(Me.lvwSpecs.SelectedItem.Key, 2)
            If gstrNodeNo <> "-" Then
                strStationNo = Me.lvwSpecs.SelectedItem.SubItems(Me.lvwSpecs.ColumnHeaders("站点").Index - 1)
            End If
        End If
        bln编辑 = (InStr(1, mstrPrivs, "存储库房") <> 0)
        Call .ShowMe(Me, lng药品ID, int用途分类, bln编辑, strStationNo)
    End With
End Sub

Private Sub mnuEditItemTabu_Click()
    With frmMediTabu
        Dim strType As String
        If Me.tvwClass.Tag = 4 Or Me.tvwClass.Tag = 3 Then '处理过滤结果，和抗菌药物
            If Me.lvwItems.SelectedItem Is Nothing Then     '如果没有记录就退出，因为无法判断药品材质
                Exit Sub
            End If
            strType = Me.lvwItems.SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1)
            strType = Switch(strType = "1", "5", strType = "2", "6", strType = "3", "7")
            .Tag = strType
        Else
            .Tag = Switch(Me.tvwClass.Tag = "0", "5", Me.tvwClass.Tag = "1", "6", Me.tvwClass.Tag = "2", "7")
        End If
        If InStr(1, mstrPrivs, "配伍禁忌关系") = 0 Then
            .cmdClose.Tag = "查阅"
        Else
            .cmdClose.Tag = "修改"
        End If
        If Me.lvwItems.SelectedItem Is Nothing Then
            .lblMedi.Tag = 0
        Else
            .lblMedi.Tag = Mid(Me.lvwItems.SelectedItem.Key, 2)
        End If
        .Show 1, Me
    End With
End Sub

Private Sub mnuEditItemUsage_Click()
    If Me.ActiveControl Is lvwItems Then
        If Me.lvwItems.SelectedItem Is Nothing Then
            If InStr(1, mstrPrivs, "用法用量") = 0 Then Exit Sub
            Call frmMediUsage.ShowMe(Me, True)
        Else
            If InStr(1, mstrPrivs, "用法用量") = 0 Then
                Call frmMediUsage.ShowMe(Me, False, Mid(Me.lvwItems.SelectedItem.Key, 2))
            Else
                If Right(Me.lvwItems.SelectedItem.Icon, 1) = "S" Then MsgBox "停用药品，不能设置用法用量！", vbExclamation, gstrSysName: Exit Sub
                Call frmMediUsage.ShowMe(Me, True, Mid(Me.lvwItems.SelectedItem.Key, 2))
                Call zlRefRecords(Mid(Me.lvwItems.SelectedItem.Key, 2))
            End If
        End If
    ElseIf Me.ActiveControl Is lvwSpecs Then
        If Me.lvwSpecs.SelectedItem.Icon = "规格S" Then MsgBox "停用规格，不能设置用法用量！", vbExclamation, gstrSysName: Exit Sub
        Call frmMediUsage.ShowMe(Me, True, Mid(Me.lvwItems.SelectedItem.Key, 2), Mid(Me.lvwSpecs.SelectedItem.Key, 2))
    End If
End Sub

Private Sub mnuEditSpecAdd_Click()
    
    If Me.lvwItems.SelectedItem Is Nothing Then MsgBox "尚未设置品种,不能增加规格！", vbExclamation, gstrSysName: Exit Sub
    mStrItem = lvwItems.SelectedItem.Key
    
    If Val(Me.tvwClass.Tag) = 2 Or mstrType = "7" Then
        With frmMediHerbalSpec
            .stbSpec.Tag = "增加"
            .mlng分类id = Val(Mid(tvwClass.SelectedItem.Key, 2))
            .lng药名id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            If Me.lvwSpecs.SelectedItem Is Nothing Then
                .lng药品ID = 0
            Else
                .lng药品ID = Mid(Me.lvwSpecs.SelectedItem.Key, 2)
            End If
            .strPrivs = Me.mstrPrivs
            .Show 1, Me
        End With
    Else
        With frmMediSpec
            .stbSpec.Tag = "增加"
            If Me.tvwClass.Tag < 3 Then
                .mlng分类id = Val(Mid(tvwClass.SelectedItem.Key, 2))
                .lng药名id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            Else
                If lvwSpecs.ListItems.Count = 0 Then '表示没有规格
                    .mlng分类id = Get分类id(Mid(lvwItems.SelectedItem.Key, 2), True)
                    .lng药名id = Mid(Me.lvwItems.SelectedItem.Key, 2)
                Else '表示有规格
                    .mlng分类id = Get分类id(Mid(lvwSpecs.SelectedItem.Key, 2))
                    .lng药名id = Mid(Me.lvwItems.SelectedItem.Key, 2)
                End If
            End If
            If Me.lvwSpecs.SelectedItem Is Nothing Then
                .lng药品ID = 0
            Else
                .lng药品ID = Mid(Me.lvwSpecs.SelectedItem.Key, 2)
            End If
            .strPrivs = Me.mstrPrivs
            .Show 1, Me
        End With
    End If
    Call zlRefRecords
'    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Function Get分类id(ByVal ID As Long, Optional ByVal bln品种 As Boolean) As Long
    '功能:获取药品所对应的分类
    '参数:bln品种表示是否是传入的品种id
    '返回:分类id
    Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHandle:
    If bln品种 = False Then
        gstrSql = "select c.分类id from 收费项目目录 a,药品规格 b,诊疗项目目录 c where a.id=b.药品id and b.药名id=c.id and a.id=[1]"
    Else
        gstrSql = "select 分类id from 诊疗项目目录 where id=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "查询分类id", ID)
    
    If rsTemp.RecordCount > 0 Then
        Get分类id = rsTemp!分类id
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mnuEditSpecDel_Click()
    Dim intCol As Integer
    With Me.lvwSpecs
        If .SelectedItem Is Nothing Then Exit Sub
        strTemp = Me.lvwItems.SelectedItem.Text & " " & .SelectedItem.Text & " " & .SelectedItem.SubItems(.ColumnHeaders("厂牌").Index - 1)
        If MsgBox("真的删除“" & strTemp & "”吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSql = "zl_成药规格_DELETE(" & Mid(.SelectedItem.Key, 2) & ")"
        err = 0: On Error GoTo ErrHand
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        
        '同步删除物流平台药品信息
        If Not gobjLogisticPlatform Is Nothing Then
            gobjLogisticPlatform.ClearDrugInfo Mid(.SelectedItem.Key, 2), 0
        End If
        
        Call .ListItems.Remove(.SelectedItem.Key)
    End With
    
    With Me.hgdPrice
        .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
        For intCol = 0 To .Cols - 1
            .TextMatrix(.FixedRows, intCol) = ""
        Next
    End With
    Call lvwItems_ItemClick(lvwItems.SelectedItem)
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditSpecLimit_Click()
    With frmMediLimit
        Dim strType As String
        If Me.tvwClass.Tag = 4 Or Me.tvwClass.Tag = 3 Then '单独处理过滤结果
            If Me.lvwItems.SelectedItem Is Nothing Then     '如果没有记录就退出，因为无法判断药品材质
                Exit Sub
            End If
            strType = Me.lvwItems.SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1)
            strType = Switch(strType = "1", "5", strType = "2", "6", strType = "3", "7")
            'strType = Switch(strType = "1", "5", strType = "2", "6", strType = "3", "6")
            .Tag = strType
        Else
            .Tag = Switch(Me.tvwClass.Tag = "0", "5", Me.tvwClass.Tag = "1", "6", Me.tvwClass.Tag = "2", "7")
            '.Tag = Switch(Me.tvwClass.Tag = "0", "5", Me.tvwClass.Tag = "1", "6", Me.tvwClass.Tag = "2", "6")
        End If
        .strPrivs = Me.mstrPrivs
        .Show 1, Me
    End With
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuEditSpecMod_Click()
    Dim lng药品ID As Long
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    If Me.lvwSpecs.SelectedItem Is Nothing Then Exit Sub
    If Me.lvwSpecs.SelectedItem.Icon = "规格S" Or Me.lvwSpecs.SelectedItem.Icon = "草规S" Then
        MsgBox "不能对停用药品药品进行修改！", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If Val(Me.tvwClass.Tag) = 2 Or (Val(Me.tvwClass.Tag) = 4 And mstrType = "7") Then
        With frmMediHerbalSpec
            .stbSpec.Tag = "修改"
            .lng药名id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            .lng药品ID = Mid(Me.lvwSpecs.SelectedItem.Key, 2)
            .strPrivs = Me.mstrPrivs
            lng药品ID = .lng药品ID
            .Show 1, Me
        End With
    Else
        With frmMediSpec
            .stbSpec.Tag = "修改"
            .lng药名id = Mid(Me.lvwItems.SelectedItem.Key, 2)
            .lng药品ID = Mid(Me.lvwSpecs.SelectedItem.Key, 2)
            .strPrivs = Me.mstrPrivs
            lng药品ID = .lng药品ID
            .Show 1, Me
        End With
    End If
    
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    '定位修改的药品
    On Error Resume Next
    err = 0
    Set lvwSpecs.SelectedItem = lvwSpecs.ListItems("_" & lng药品ID)
    If err <> 0 Then Set lvwSpecs.SelectedItem = lvwSpecs.ListItems(1)
End Sub

Private Sub mnuEditSpecProtocol_Click()
    With frmMediMember
        Dim strType As String
        If Me.tvwClass.Tag = 4 Or Me.tvwClass.Tag = 3 Then '单独处理过滤结果
            If Me.lvwItems.SelectedItem Is Nothing Then     '如果没有记录就退出，因为无法判断药品材质
                Exit Sub
            End If
            strType = Me.lvwItems.SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1)
            strType = Switch(strType = "1", "5", strType = "2", "6", strType = "3", "7")
            .Tag = strType
        Else
            .Tag = Switch(Me.tvwClass.Tag = "0", "5", Me.tvwClass.Tag = "1", "6", Me.tvwClass.Tag = "2", "7")
        End If
        If InStr(1, mstrPrivs, "协定药品构成") = 0 Then
            .cmdClose.Tag = "查阅"
        Else
            .cmdClose.Tag = "修改"
        End If
        If Me.lvwSpecs.SelectedItem Is Nothing Then
            .lblMedi.Tag = 0
        Else
            .lblMedi.Tag = Mid(Me.lvwSpecs.SelectedItem.Key, 2)
        End If
        .msfMember.Tag = "协定"
        .Show 1, Me
    End With
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuEditSpecSelf_Click()
    With frmMediMember
        Dim strType As String
        If Me.tvwClass.Tag = 4 Or Me.tvwClass.Tag = 3 Then '单独处理过滤结果
            If Me.lvwItems.SelectedItem Is Nothing Then     '如果没有记录就退出，因为无法判断药品材质
                Exit Sub
            End If
            strType = Me.lvwItems.SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1)
            strType = Switch(strType = "1", "5", strType = "2", "6", strType = "3", "7")
            .Tag = strType
        Else
            .Tag = Switch(Me.tvwClass.Tag = "0", "5", Me.tvwClass.Tag = "1", "6", Me.tvwClass.Tag = "2", "7")
        End If
        If InStr(1, mstrPrivs, "自制药品构成") = 0 Then
            .cmdClose.Tag = "查阅"
        Else
            .cmdClose.Tag = "修改"
        End If
        If Me.lvwSpecs.SelectedItem Is Nothing Then
            .lblMedi.Tag = 0
        Else
            .lblMedi.Tag = Mid(Me.lvwSpecs.SelectedItem.Key, 2)
        End If
        .msfMember.Tag = "自制"
        .Show 1, Me
    End With
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuEditSpecUnit_Click()
    On Error Resume Next
    '招标药品中标单位设置
    With frmMediUnit
        Dim strType As String
        If Me.tvwClass.Tag = 4 Or Me.tvwClass.Tag = 3 Then '单独处理过滤结果
            If Me.lvwItems.SelectedItem Is Nothing Then     '如果没有记录就退出，因为无法判断药品材质
                Exit Sub
            End If
            strType = Me.lvwItems.SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1)
            strType = Switch(strType = "1", "5", strType = "2", "6", strType = "3", "7")
            .frmTag = strType
        Else
            .frmTag = Switch(Me.tvwClass.Tag = "0", "5", Me.tvwClass.Tag = "1", "6", Me.tvwClass.Tag = "2", "7")
        End If
        If Me.lvwSpecs.SelectedItem Is Nothing Then
            .lblTag = 0
        Else
            .lblTag = Mid(Me.lvwSpecs.SelectedItem.Key, 2)
        End If
        .strPrivs = Me.mstrPrivs
        .Show 1, Me
    End With
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuEditStart_Click()
    If Me.ActiveControl.Name = Me.lvwItems.Name Then
        With Me.lvwItems
            If .SelectedItem Is Nothing Then Exit Sub
            If .SelectedItem.Icon = "成药U" Or .SelectedItem.Icon = "草药U" Then Exit Sub
            
            If MsgBox("真的重新启用“" & .SelectedItem.Text & "”吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
'            If Val(Me.tvwClass.Tag) < 2 Or (Val(Me.tvwClass.Tag) = 3 And mstrType <> "7") Then
                gstrSql = "zl_成药品种_REUSE(" & Mid(.SelectedItem.Key, 2) & ")"
'            ElseIf Val(Me.tvwClass.Tag) = 2 Or (Val(Me.tvwClass.Tag) = 3 And mstrType = "7") Then
'                gstrSql = "zl_草药药品_REUSE(" & Mid(.SelectedItem.Key, 2) & ")"
'            End If
            err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            If Val(Me.tvwClass.Tag) < 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType <> "7" Then
                .SelectedItem.Icon = "成药U": .SelectedItem.SmallIcon = "成药U"
            ElseIf Val(Me.tvwClass.Tag) = 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType = "7" Then
                .SelectedItem.Icon = "草药U": .SelectedItem.SmallIcon = "草药U"
            End If
            '恢复启用项目显示颜色
            .SelectedItem.ForeColor = .ForeColor
            For intCount = 1 To .ColumnHeaders.Count - 1
                .SelectedItem.ListSubItems(intCount).ForeColor = .ForeColor
            Next
        End With
    Else
        With Me.lvwSpecs
            If .Visible = False Then Exit Sub
            If .SelectedItem Is Nothing Then Exit Sub
            If .SelectedItem.Icon = "规格U" Then Exit Sub
            
            strTemp = Me.lvwItems.SelectedItem.Text & " " & .SelectedItem.Text & " " & .SelectedItem.SubItems(.ColumnHeaders("厂牌").Index - 1)
            If MsgBox("真的重新启用“" & strTemp & "”吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            gstrSql = "zl_成药规格_REUSE(" & Mid(.SelectedItem.Key, 2) & ")"
            err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            If Val(Me.tvwClass.Tag) < 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType <> "7" Then
                .SelectedItem.Icon = "规格U": .SelectedItem.SmallIcon = "规格U"
            Else
                .SelectedItem.Icon = "草规U": .SelectedItem.SmallIcon = "草规U"
            End If
            '恢复启用项目显示颜色
            .SelectedItem.ForeColor = .ForeColor
            For intCount = 1 To .ColumnHeaders.Count - 1
                .SelectedItem.ListSubItems(intCount).ForeColor = .ForeColor
            Next
        End With
    End If
    
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditStop_Click()
    Dim lng药品ID As Long
    Dim rsTemp As ADODB.Recordset
    Dim blnStop As Boolean
    
    If Me.ActiveControl.Name = Me.lvwItems.Name Then
        With Me.lvwItems
            If .SelectedItem Is Nothing Then Exit Sub
            If .SelectedItem.Icon = "成药S" Or .SelectedItem.Icon = "草药S" Then Exit Sub
            
            gstrSql = "select b.实际数量 from 药品规格 a,药品库存 b where a.药品id=b.药品id and a.药名id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "库存检查", Mid(.SelectedItem.Key, 2))
            
            If rsTemp.RecordCount > 0 Then
                If IIf(IsNull(rsTemp!实际数量), "0", rsTemp!实际数量) > 0 Then
                    If MsgBox("该药品有库存数量，确定停用？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                    blnStop = True
                End If
            End If
            
            If blnStop = False Then
                If MsgBox("真的要停用“" & .SelectedItem.Text & "”吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            End If
            
'            If Val(Me.tvwClass.Tag) < 2 Or (Val(Me.tvwClass.Tag) = 3 And mstrType <> "7") Then
                gstrSql = "zl_成药品种_STOP(" & Mid(.SelectedItem.Key, 2) & ")"
'            ElseIf Val(Me.tvwClass.Tag) = 2 Or (Val(Me.tvwClass.Tag) = 3 And mstrType = "7") Then
'                gstrSql = "zl_草药药品_STOP(" & Mid(.SelectedItem.Key, 2) & ")"
'            End If
            err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            If Me.mnuViewStoped.Checked = True Then
                If Val(Me.tvwClass.Tag) < 2 Or (Val(Me.tvwClass.Tag) = 4 And mstrType <> "7") Then
                    .SelectedItem.Icon = "成药S": .SelectedItem.SmallIcon = "成药S"
                ElseIf Val(Me.tvwClass.Tag) = 2 Or (Val(Me.tvwClass.Tag) = 4 And mstrType = "7") Then
                    .SelectedItem.Icon = "草药S": .SelectedItem.SmallIcon = "草药S"
                End If
                '将停用项目显示为红色
                .SelectedItem.ForeColor = mconColor_Stop
                For intCount = 1 To .ColumnHeaders.Count - 1
                    .SelectedItem.ListSubItems(intCount).ForeColor = mconColor_Stop
                Next
            Else
                Call .ListItems.Remove(.SelectedItem.Key)
            End If
        End With
    Else
        With Me.lvwSpecs
            If .Visible = False Then Exit Sub
            If .SelectedItem Is Nothing Then Exit Sub
            If .SelectedItem.Icon = "规格S" Then Exit Sub
            
            gstrSql = "select 实际数量 from 药品库存 where 药品id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "库存检查", Mid(.SelectedItem.Key, 2))
            
            If rsTemp.RecordCount > 0 Then
                If IIf(IsNull(rsTemp!实际数量), "0", rsTemp!实际数量) > 0 Then
                    If MsgBox("该药品有库存数量，确定停用？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                    blnStop = True
                End If
            End If
            
            strTemp = Me.lvwItems.SelectedItem.Text & " " & .SelectedItem.Text & " " & .SelectedItem.SubItems(.ColumnHeaders("厂牌").Index - 1)
            
            If blnStop = False Then
                If MsgBox("真的要停用“" & .SelectedItem.Text & "”吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            End If
            
            gstrSql = "zl_成药规格_STOP(" & Mid(.SelectedItem.Key, 2) & ")"
            err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            If Me.mnuViewStoped.Checked = True Then
                If Val(Me.tvwClass.Tag) < 2 Or (Val(Me.tvwClass.Tag) = 4 And mstrType <> "7") Then
                    .SelectedItem.Icon = "规格S": .SelectedItem.SmallIcon = "规格S"
                Else
                    .SelectedItem.Icon = "草规S": .SelectedItem.SmallIcon = "草规S"
                End If
                '将停用项目显示为红色
                .SelectedItem.ForeColor = mconColor_Stop
                For intCount = 1 To .ColumnHeaders.Count - 1
                    .SelectedItem.ListSubItems(intCount).ForeColor = mconColor_Stop
                Next
            Else
                Call .ListItems.Remove(.SelectedItem.Key)
            End If
        End With
    End If
    
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuFileExcel_Click()
    Call zlRptPrint(3)
End Sub

Private Sub mnuFilePara_Click()
    '模块公共参数已经调整到药品参数设置模块，目前没有私有或本机参数，暂时屏蔽参数设置界面
'    frmMediPara.ShowMe mstrPrivs, Me
End Sub

Private Sub mnuFilePreview_Click()
    Call zlRptPrint(0)
End Sub

Private Sub mnuFilePrint_Click()
    Call zlRptPrint(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuFilter_Click()
    With frmMediFilter
        Call .ShowMe(Me, mnuViewStoped.Checked, mbln自管药)
     End With
End Sub

Private Sub mnuhelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuPriceChargeSet1_Click()
    Call mnuPriceChargeSet_Click
End Sub

Private Sub mnuPriceLists_Click()
    Dim str类别 As String
    Dim lng分类id As Long
    
    Select Case Val(Me.tvwClass.Tag)
    Case 0
        str类别 = "5"
    Case 1
        str类别 = "6"
    Case 2
        str类别 = "7"
    End Select
    
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    lng分类id = Val(Mid(Me.tvwClass.SelectedItem.Key, 2))
    
    Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1023_2", "ZL8_BILL_1023_2"), Me, "类别=" & str类别, "分类=" & lng分类id)
End Sub


Private Sub mnuPriceTable_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then
        Exit Sub
    End If
    Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "Zl1_BILL_1023_1", "ZL8_BILL_1023_1"), Me)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    '默认参数：分类=分类id，品种=药名id，规格=药品id
    Dim lng分类id As Long
    Dim lng药名id As Long
    Dim lng规格id As Long
    
    If Me.tvwClass.Tag <> 3 Then
        If Not Me.tvwClass.SelectedItem Is Nothing Then
            lng分类id = Val(Mid(Me.tvwClass.SelectedItem.Key, 2))
        End If
    End If
    
    If Not Me.lvwItems.SelectedItem Is Nothing Then
        lng药名id = Mid(lvwItems.SelectedItem.Key, 2)
    End If
    
    If Not Me.lvwSpecs.SelectedItem Is Nothing Then
        lng规格id = Mid(lvwSpecs.SelectedItem.Key, 2)
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "分类=" & IIf(lng分类id = 0, "", lng分类id), _
        "品种=" & IIf(lng药名id = 0, "", lng药名id), _
        "规格=" & IIf(lng规格id = 0, "", lng规格id))
End Sub

Private Sub mnuUploadDrugInfo_Click()
    '批量上传药品信息
    If Not gobjLogisticPlatform Is Nothing Then
        gobjLogisticPlatform.UploadDrugInfo Me, gcnOracle, 0
    End If
End Sub

Private Sub mnuViewFind_Click()
    With frmMediFind
        Call .ShowMe(Me, mnuViewStoped.Checked, mbln自管药)
    End With
End Sub

Private Sub mnuViewFindNext_Click()
    On Error Resume Next
    
    Select Case Val(tvwClass.Tag)
    Case 0
        frmMediFind.Tag = 5: Me.Caption = "西成药查找..."
    Case 1
        frmMediFind.Tag = 6: Me.Caption = "中成药查找..."
    Case 2
        frmMediFind.Tag = 7: Me.Caption = "中草药查找..."
    End Select
    Call frmMediFind.FindNext
End Sub

Private Sub mnuViewList_Click()
    mstrFindValue = ""
    Set mrsFind = Nothing
    If Me.tvwClass.SelectedItem Is Nothing Then
        Exit Sub
    End If
    Me.mnuViewList.Checked = Not Me.mnuViewList.Checked
    Call zlRefClasses
End Sub
Private Sub mnuViewPrices_Click()
    Me.mnuViewPrices.Checked = Not Me.mnuViewPrices.Checked
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuViewRefer_Click()
    Call gobjKernel.InitCISKernel(gcnOracle, Me, glngSys, mstrPrivs)
    If Me.lvwItems.SelectedItem Is Nothing Then
        Call gobjKernel.ShowClincHelp(0, Me)
    Else
        Call gobjKernel.ShowClincHelp(0, Me, Val(Mid(Me.lvwItems.SelectedItem.Key, 2)))
    End If
End Sub

Private Sub mnuViewRefresh_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Call zlRefRecords
End Sub

Private Sub mnuViewShowAll_Click()
    On Error GoTo ErrHandle
    mnuViewShowAll.Checked = Not mnuViewShowAll.Checked
    If tvwClass.SelectedItem Is Nothing Then
        If tvwClass.Nodes.Count > 0 Then
            MsgBox "请选择一下分类！", vbInformation, gstrSysName
        Else
            MsgBox "无任何分类可显示！", vbInformation, gstrSysName
        End If
        Exit Sub
    End If
    Call zlRefRecords
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuViewStates_Click()
    Me.mnuViewStates.Checked = Not Me.mnuViewStates.Checked
    Me.stbThis.Visible = Me.mnuViewStates.Checked
    Form_Resize
End Sub

Private Sub mnuViewStoped_Click()
    mstrFindValue = ""
    Set mrsFind = Nothing
    If Me.tvwClass.SelectedItem Is Nothing Then
        Exit Sub
    End If
    Me.mnuViewStoped.Checked = Not Me.mnuViewStoped.Checked
    Call zlRefRecords
End Sub

Private Sub mnuViewToolbarStand_Click()
    Me.mnuViewToolbarStand.Checked = Not Me.mnuViewToolbarStand.Checked
    Me.clbThis.Visible = Me.mnuViewToolbarStand.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolBarText_Click()
    Dim i As Integer
    Me.mnuViewToolbarText.Checked = Not Me.mnuViewToolbarText.Checked
    If Me.mnuViewToolbarText.Checked Then
        For i = 1 To Me.tlbThis.Buttons.Count
            Me.tlbThis.Buttons(i).Caption = Me.tlbThis.Buttons(i).Tag
        Next
    Else
        For i = 1 To Me.tlbThis.Buttons.Count
            Me.tlbThis.Buttons(i).Caption = ""
        Next
    End If
    Me.clbThis.Bands(1).MinHeight = Me.tlbThis.Height
    Me.clbThis.Refresh
    Form_Resize
End Sub

Private Sub picClass_Resize()
    Dim intCount As Integer
    err = 0: On Error Resume Next
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        Me.cmdKind(intCount).Left = Me.picClass.ScaleLeft + 15
        Me.cmdKind(intCount).Width = Me.picClass.ScaleWidth
        Me.cmdKind(intCount).Height = 300
        
        If intCount = 2 And mbln自管药 = True Then '自管药不显示中草药 单独处理
            cmdKind(intCount).Visible = False
        End If
        If intCount <= mintIndex Then
            If mbln自管药 = True And intCount > 2 Then
                Me.cmdKind(intCount).Top = Me.picClass.ScaleTop + 285 * (intCount - 1)
                Me.tvwClass.Top = Me.picClass.ScaleTop + 285 * intCount
            Else
                Me.cmdKind(intCount).Top = Me.picClass.ScaleTop + 285 * intCount
                Me.tvwClass.Top = Me.picClass.ScaleTop + 285 * (intCount + 1)
            End If
        Else
            If mbln自管药 = True And intCount > 2 Then
                Me.cmdKind(intCount).Top = Me.picClass.ScaleHeight - 285 * (Me.cmdKind.UBound - intCount + 2)
            Else
                Me.cmdKind(intCount).Top = Me.picClass.ScaleHeight - 285 * (Me.cmdKind.UBound - intCount + 1)
            End If
        End If
    Next
    Me.tvwClass.Left = Me.picClass.ScaleLeft + 15
    Me.tvwClass.Width = Me.picClass.ScaleWidth
    Me.tvwClass.Height = Me.picClass.ScaleHeight - 285 * (Me.cmdKind.UBound + 1) - 15
End Sub

Private Sub picHBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Me.picHBar.Top = Me.picHBar.Top + y
    End If
End Sub

Private Sub picHBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Call Form_Resize
    End If
End Sub

Private Sub picVBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Me.picVBar.Left = Me.picVBar.Left + x
    End If
End Sub

Private Sub picVBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Call Form_Resize
    End If
End Sub

Private Sub tabContent_Click(PreviousTab As Integer)
    mintPage = Me.tabContent.Tab
    Select Case Me.tabContent.Tab
    Case 0
        Me.lvwSpecs.Visible = True
        Me.fraComment(0).Visible = True
        Me.hgdPrice.Visible = False
        Me.fraComment(1).Visible = False
        Me.hgdCost.Visible = False
    Case 1
        Me.lvwSpecs.Visible = False
        Me.fraComment(0).Visible = False
        Me.hgdPrice.Visible = True
        Me.fraComment(1).Visible = True
        Me.hgdCost.Visible = False
    Case 2
        Me.lvwSpecs.Visible = False
        Me.fraComment(0).Visible = False
        Me.hgdPrice.Visible = False
        Me.fraComment(1).Visible = False
        Me.hgdCost.Visible = True
        hgdCharge.Visible = False
        Call GetCostAdjust(mlngCurrDrug)
    Case 3
        Me.lvwSpecs.Visible = False
        Me.fraComment(0).Visible = False
        Me.hgdPrice.Visible = False
        Me.fraComment(1).Visible = False
        hgdCost.Visible = False
        Me.hgdCharge.Visible = True
        Call GetChargeSet(mlngCurrDrug)
    End Select
End Sub

Private Sub tlbThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Preview"
        Call mnuFilePreview_Click
    Case "Print"
        Call mnuFilePrint_Click
    Case "Class"
        Call PopupMenu(Me.mnuClass, 2)
    Case "Item"
        Call zlPopupEditMenu(1, False)
    Case "Spec"
        Call zlPopupEditMenu(2, False)
    Case "Start"
        If Me.ActiveControl Is tvwClass Then
            Call mnuClassStar_Click
        Else
            Call mnuEditStart_Click
        End If
    Case "Stop"
        If Me.ActiveControl Is tvwClass Then
            Call mnuClassStop_Click
        Else
            Call mnuEditStop_Click
        End If
    Case "Limit"
        Call mnuEditSpecLimit_Click
    Case "Find"
        Call mnuViewFind_Click
    Case "Help"
        Call mnuHelpHelp_Click
    Case "Exit"
        Call mnuFileExit_Click
    Case "Filter"
        Call mnuFilter_Click
    End Select
End Sub

Private Sub tlbThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    PopupMenu Me.mnuViewToolbar, 2
End Sub

Private Sub tvwClass_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If Val(Me.tvwClass.Tag) >= 3 Then           '=3是显示过滤结果的状态，不允许编辑类别
        Exit Sub
    End If
    If mbln自管药 = False Then
        Call zlPopupClassMenu
    End If
End Sub

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim bln启用 As Boolean
    Dim bln停用 As Boolean
    Dim objItem As ListItem
    Dim strTemp As String
    Dim strNode As String
    
    If mstrKey <> Node.Key Then
        mstrKey = Node.Key
    Else
        Exit Sub
    End If
    
    If InStr(1, mstrPrivs, ";药品启用;") = 0 Then
        mnuClassStar.Visible = False
        tlbThis.Buttons("Start").Visible = False
    Else
        mnuClassStar.Visible = True
        tlbThis.Buttons("Start").Visible = True
    End If
    If InStr(1, mstrPrivs, ";药品停用;") = 0 Then
        mnuClassStop.Visible = False
        tlbThis.Buttons("Stop").Visible = False
    Else
        mnuClassStop.Visible = True
        tlbThis.Buttons("Stop").Visible = True
    End If
    If tlbThis.Buttons("Start").Visible = False Or tlbThis.Buttons("Stop").Visible = False Then
        tlbThis.Buttons("sp2").Visible = False
    End If
    
    Call zlRefRecords
    
    bln启用 = mnuClassStar.Visible
    bln停用 = mnuClassStop.Visible
    
    If mstr分类 = "0" Then
        mstrNodeSelect西成药 = tvwClass.SelectedItem.Key
        strTemp = mstrItemClick西成药
        strNode = mstrNodeClick西成药
    ElseIf mstr分类 = "1" Then
        mstrNodeSelect中成药 = tvwClass.SelectedItem.Key
        strTemp = mstrItemClick中成药
        strNode = mstrNodeClick中成药
    ElseIf mstr分类 = "2" Then
        mstrNodeSelect中草药 = tvwClass.SelectedItem.Key
        strTemp = mstrItemClick中草药
        strNode = mstrNodeClick中草药
    End If
    
    If strNode = Node.Key Then    '定位药品
        For Each objItem In lvwItems.ListItems
            If objItem.Key = strTemp Then
                lvwItems.ListItems(objItem.Key).Selected = True
                Exit For
            End If
        Next
    End If
        
    With tvwClass
        If .Nodes.Count = 0 Then Exit Sub
        If lvwItems.SelectedItem Is Nothing Then
            mnuEditItemMod.Enabled = False
            mnuEditItemDel.Enabled = False
            mnuEditItemBill.Enabled = False
            mnuEditStart.Enabled = False
            mnuEditStop.Enabled = False
            tlbThis.Buttons("Start").Enabled = False
            tlbThis.Buttons("Stop").Enabled = False
            mnuEditSpecAdd.Enabled = False
            mnuEditSpecMod.Enabled = False
            mnuEditSpecDel = False
        Else
            mnuEditItemMod.Enabled = True
            mnuEditItemDel.Enabled = True
            mnuEditItemBill.Enabled = True
            mnuEditStart.Enabled = bln启用
            mnuEditStop.Enabled = bln停用
            tlbThis.Buttons("Start").Enabled = bln启用
            tlbThis.Buttons("Stop").Enabled = bln停用
            If lvwItems.SelectedItem.ForeColor = vbRed Then
                mnuEditSpecAdd.Enabled = False
                mnuEditSpecMod.Enabled = False
                mnuEditSpecDel = True
            Else
                mnuEditSpecAdd.Enabled = True
                mnuEditSpecMod.Enabled = True
                mnuEditSpecDel = True
            End If
            If lvwSpecs.SelectedItem Is Nothing Then
                mnuEditSpecMod.Enabled = False
                mnuEditSpecDel = False
            End If
        End If
    End With
End Sub

Private Sub zlRefPurview()
    '---------------------------------------------
    '填写诊疗分类项目(此处为药品分类)并按照不同类型调整界面
    '---------------------------------------------
    '材质权限控制
    Me.mnuEdit.Enabled = True
    Me.mnuPrice.Enabled = True
    If Val(Me.tvwClass.Tag) = 0 Then
        If InStr(1, mstrPrivs, "管理西成药") = 0 Then
            mnuClassAdd.Visible = False
            mnuClassDel.Visible = False
            mnuClassMod.Visible = False
        Else
            mnuClass.Visible = True
            tlbThis.Buttons("Class").Visible = True
            mnuClassAdd.Visible = True
            mnuClassDel.Visible = True
            mnuClassMod.Visible = True
        End If
        If InStr(1, mstrPrivs, "管理西成药品种") = 0 Then
            mnuEditItemAdd.Visible = False
            mnuEditItemDel.Visible = False
            mnuEditItemMod.Visible = False
        Else
            mnuEditItemAdd.Visible = True
            mnuEditItemDel.Visible = True
            mnuEditItemMod.Visible = True
        End If
        If InStr(1, mstrPrivs, "管理西成药规格") = 0 Then
            mnuEditSpecAdd.Visible = False
            mnuEditSpecDel.Visible = False
            mnuEditSpecMod.Visible = False
        Else
            mnuEditSpecAdd.Visible = True
            mnuEditSpecDel.Visible = True
            mnuEditSpecMod.Visible = True
        End If
    End If
    If Val(Me.tvwClass.Tag) = 1 Then
        If InStr(1, mstrPrivs, "管理中成药") = 0 Then
            mnuClassAdd.Visible = False
            mnuClassDel.Visible = False
            mnuClassMod.Visible = False
        Else
            mnuClass.Visible = True
            tlbThis.Buttons("Class").Visible = True
            mnuClassAdd.Visible = True
            mnuClassDel.Visible = True
            mnuClassMod.Visible = True
        End If
        If InStr(1, mstrPrivs, "管理中成药品种") = 0 Then
            mnuEditItemAdd.Visible = False
            mnuEditItemDel.Visible = False
            mnuEditItemMod.Visible = False
        Else
            mnuEditItemAdd.Visible = True
            mnuEditItemDel.Visible = True
            mnuEditItemMod.Visible = True
        End If
        If InStr(1, mstrPrivs, "管理中成药规格") = 0 Then
            mnuEditSpecAdd.Visible = False
            mnuEditSpecDel.Visible = False
            mnuEditSpecMod.Visible = False
        Else
            mnuEditSpecAdd.Visible = True
            mnuEditSpecDel.Visible = True
            mnuEditSpecMod.Visible = True
        End If
    End If
    If Val(Me.tvwClass.Tag) = 2 Then
        If InStr(1, mstrPrivs, "管理中草药") = 0 Then
            mnuClassAdd.Visible = False
            mnuClassDel.Visible = False
            mnuClassMod.Visible = False
        Else
            mnuClass.Visible = True
            tlbThis.Buttons("Class").Visible = True
            mnuClassAdd.Visible = True
            mnuClassDel.Visible = True
            mnuClassMod.Visible = True
        End If
        If InStr(1, mstrPrivs, "管理中草药品种") = 0 Then
            mnuEditItemAdd.Visible = False
            mnuEditItemDel.Visible = False
            mnuEditItemMod.Visible = False
        Else
            mnuEditItemAdd.Visible = True
            mnuEditItemDel.Visible = True
            mnuEditItemMod.Visible = True
        End If
        If InStr(1, mstrPrivs, "管理中草药规格") = 0 Then
            mnuEditSpecAdd.Visible = False
            mnuEditSpecDel.Visible = False
            mnuEditSpecMod.Visible = False
        Else
            mnuEditSpecAdd.Visible = True
            mnuEditSpecDel.Visible = True
            mnuEditSpecMod.Visible = True
        End If
    End If
    If Val(Me.tvwClass.Tag) = 4 And InStr(1, mstrPrivs, "管理西成药") = 0 _
        And InStr(1, mstrPrivs, "管理中成药") = 0 _
        And InStr(1, mstrPrivs, "管理中草药") = 0 Then
        Me.mnuEdit.Enabled = False: Me.mnuPrice.Enabled = False
    End If
    
    If InStr(1, mstrPrivs, ";药品启用;") = 0 Then
        mnuClassStar.Visible = False
        mnuEditStart.Visible = False
        tlbThis.Buttons("Start").Visible = False
    Else
        mnuClassStar.Visible = True
        mnuEditStart.Visible = True
        tlbThis.Buttons("Start").Visible = True
    End If
    If InStr(1, mstrPrivs, ";药品停用;") = 0 Then
        mnuClassStop.Visible = False
        mnuEditStop.Visible = False
        tlbThis.Buttons("Stop").Visible = False
    Else
        mnuClassStop.Visible = True
        mnuEditStop.Visible = True
        tlbThis.Buttons("Stop").Visible = True
    End If
    If tlbThis.Buttons("Start").Visible = False Or tlbThis.Buttons("Stop").Visible = False Then
        tlbThis.Buttons("sp2").Visible = False
    End If
    
    If Val(Me.tvwClass.Tag) = 2 Then
        Me.mnuEditItemUsage.Visible = False
    Else
        Me.mnuEditItemUsage.Visible = True
    End If
    
    If InStr(1, mstrPrivs, "对应处方") = 0 Then
        Me.mnuEditItemBill.Enabled = False
    Else
        Me.mnuEditItemBill.Enabled = Me.mnuEdit.Enabled
    End If
    
    If InStr(1, mstrPrivs, "售价管理") = 0 And InStr(1, mstrPrivs, "成本价管理") = 0 Then
    Else
        If InStr(mstrPrivs, "费别设置") = 0 Then
            mnuPriceChargeSet.Visible = False
        Else
            mnuPriceChargeSet.Visible = True
            mnuPriceChargeSet.Enabled = Me.mnuPrice.Enabled
        End If
    End If
    Me.tlbThis.Buttons("Limit").Enabled = Me.mnuEditSpecLimit.Enabled
    
    If InStr(1, mstrPrivs, "调价记录查询") = 0 Then
        Me.mnuPriceLists.Visible = False
        Me.mnuPriceTable.Visible = False
    Else
        Me.mnuPriceSpt1.Visible = Me.mnuPrice.Enabled
        Me.mnuPriceLists.Visible = Me.mnuPrice.Enabled
        Me.mnuPriceTable.Visible = Me.mnuPrice.Enabled
    End If
    
    '调整显示界面
    If Val(Me.tvwClass.Tag) > 2 Then '大于2的为抗菌药物和过滤结果
        Me.mnuClass.Visible = False
        Me.tlbThis.Buttons("Class").Visible = False
    End If
End Sub

Private Sub zlRefClasses(Optional lngNode As Long)
    Dim intCol As Integer
    Dim intType As Integer
    Dim i As Integer
    
    On Error GoTo ErrHand
    
    Call zlRefPurview '权限验证
    
    If Val(Me.tvwClass.Tag) < 3 Then
        '西成药、中成药、草药
'        Me.mnuEditSpecAdd.Visible = True
'        Me.mnuEditSpecMod.Visible = True
'        Me.mnuEditSpecDel.Visible = True
'        Me.tlbThis.Buttons("Spec").Visible = True
        Me.lvwItems.ListItems.Clear
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "_名称", "名称", 2500
            .Add , "_编码", "编码", 1000
            .Add , "_剂量单位", "剂量单位", 900
            .Add , "_剂型", "剂型", 1800
            '.Add , "_服务对象", "服务对象", 1100
            .Add , "_处方类型", "处方类型", 1000
            .Add , "_处方限量", "处方限量", 900
            .Add , "_过敏试验", "过敏试验", 900
            .Add , "_毒理", "毒理", 750
            .Add , "_货源", "货源", 600
            .Add , "_价值", "价值", 600
            .Add , "_梯次", "梯次", 600
            .Add , "_原料药", "原料药", 750
            If Val(Me.tvwClass.Tag) = 2 Then
                .Add , "_单味使用", "单味使用", 900
            Else
                .Add , "_急救药", "急救药", 750
                .Add , "_新药", "新药", 600
                .Add , "_原研药", "原研药", 800
                .Add , "_专利药", "专利药", 800
                .Add , "_单独定价", "单独定价", 900
                .Add , "_抗菌药物", "抗菌药物", 1500
            End If
            .Add , "_按药品下长期医嘱", "按药品下长期医嘱", 1500
            .Add , "_适用性别", "适用性别", 1200
            .Add , "_辅助用药", "辅助用药", 1200
        End With
        With Me.lvwItems
            .ColumnHeaders("_编码").Position = 1
            .SortKey = .ColumnHeaders("_编码").Index - 1
            .SortOrder = lvwAscending
        End With
        
        '规格列设置
        Me.lvwSpecs.ListItems.Clear
        With Me.lvwSpecs.ColumnHeaders
            .Clear
            .Add , "规格", "规格", 1500: .Add , "编码", "编码", 1100
            If Not (Val(Me.tvwClass.Tag) = 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType = "7") Then
                .Add , "厂牌", "生产商", 2000
                .Add , "商品名", "商品名", 2000
            Else
                .Add , "厂牌", "生产商", 2000
                .Add , "原产地", "原产地", 2000
            End If
            .Add , "服务对象", "服务对象", 1200: .Add , "医保类型", "医保类型", 900: .Add , "药品来源", "药品来源", 900: .Add , "自制", "自制", 600
            .Add , "协定", "协定", 600
            If Val(Me.tvwClass.Tag) = 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType = "7" Then .Add , "中药形态", "中药形态", 900
            .Add , "招标", "招标", 600: .Add , "批准文号", "批准文号", 1600
            .Add , "合同单位", "合同单位", 3000: .Add , "说明", "说明", 2000
            .Add , "备选码", "备选码", 1000: .Add , "站点", "院区", IIf(gstrNodeNo = "-", 0, 1000)
        End With
        With Me.lvwSpecs
            .ColumnHeaders("编码").Position = 1
            .SortKey = .ColumnHeaders("编码").Index - 1
            .SortOrder = lvwAscending
        End With
        
        Me.tabContent.TabVisible(0) = True
        Me.tabContent.Tab = mintPage
        Call tabContent_Click(mintPage)
    End If
    
   
    With Me.hgdPrice
        .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
        For intCol = 0 To .Cols - 1
            .TextMatrix(.FixedRows, intCol) = ""
        Next
    End With
    Call RestoreListViewState(Me.lvwItems, Me.Name & Val(Me.tvwClass.Tag), Me.lvwItems.View)
    
    
    ''''''''''''''''''''''''''''''如果是过滤结果，则单独处理
    If Me.tvwClass.Tag >= 3 Then
        Me.tvwClass.Nodes.Clear
        Me.lvwItems.ListItems.Clear
        Me.lvwSpecs.ListItems.Clear
       
        '按材质设置显示列
        Me.mnuEditItemAdd.Visible = True
        Me.mnuEditSpecAdd.Visible = True
        Me.mnuEditSpecMod.Visible = True
        Me.mnuEditSpecDel.Visible = True
        Me.mnuEditItemUsage.Visible = True
        Me.tlbThis.Buttons("Spec").Visible = True
        Me.lvwItems.ListItems.Clear
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "_名称", "名称", 2500
            .Add , "_编码", "编码", 1000
            .Add , "_剂量单位", "剂量单位", 900
            .Add , "_剂型", "剂型", 600
            '.Add , "_服务对象", "服务对象", 1100
            .Add , "_处方类型", "处方类型", 1000
            .Add , "_处方限量", "处方限量", 900
            .Add , "_过敏试验", "过敏试验", 900
            .Add , "_毒理", "毒理", 750
            .Add , "_货源", "货源", 600
            .Add , "_价值", "价值", 600
            .Add , "_梯次", "梯次", 600
            .Add , "_原料药", "原料药", 750
            If mstrType = "7" Then
                .Add , "_单味使用", "单味使用", 900
            Else
                .Add , "_急救药", "急救药", 750
                .Add , "_新药", "新药", 600
                .Add , "_原研药", "原研药", 800
                .Add , "_专利药", "专利药", 800
                .Add , "_单独定价", "单独定价", 900
                .Add , "_抗菌药物", "抗菌药物", 1500
            End If
            .Add , "_按药品下长期医嘱", "按药品下长期医嘱", 1500
            .Add , "_适用性别", "适用性别", 1200
            .Add , "_类型", "类型", 0
            .Add , "_分类id", "分类id", 0
            .Add , "_辅助用药", "辅助用药", 1200
        End With
        With Me.lvwItems
            .ColumnHeaders("_编码").Position = 1
            .SortKey = .ColumnHeaders("_编码").Index - 1
            .SortOrder = lvwAscending
        End With
        
        '规格列设置
        Me.lvwSpecs.ListItems.Clear
        With Me.lvwSpecs.ColumnHeaders
            .Clear
            .Add , "规格", "规格", 1500: .Add , "编码", "编码", 1100
            If Not (Val(Me.tvwClass.Tag) = 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType = "7") Then
                .Add , "厂牌", "生产商", 2000
                .Add , "商品名", "商品名", 2000
            Else
                .Add , "厂牌", "生产商", 2000
                .Add , "原产地", "原产地", 2000
            End If
            .Add , "服务对象", "服务对象", 1500: .Add , "医保类型", "医保类型", 900: .Add , "药品来源", "药品来源", 900: .Add , "自制", "自制", 600
            .Add , "协定", "协定", 600
            If Val(Me.tvwClass.Tag) = 2 Or Val(Me.tvwClass.Tag) = 4 And mstrType = "7" Then .Add , "中药形态", "中药形态", 900
            .Add , "招标", "招标", 600: .Add , "批准文号", "批准文号", 1600: .Add , "合同单位", "合同单位", 3000: .Add , "说明", "说明", 2000
            .Add , "备选码", "备选码", 1000: .Add , "站点", "院区", IIf(gstrNodeNo = "-", 0, 1000)
        End With
        With Me.lvwSpecs
            .ColumnHeaders("编码").Position = 1
            .SortKey = .ColumnHeaders("编码").Index - 1
            .SortOrder = lvwAscending
        End With
        
        Me.tabContent.TabVisible(0) = True
        Me.tabContent.Tab = mintPage
        Call tabContent_Click(mintPage)
        Call setMenu自管药
       If Val(Me.tvwClass.Tag) = 4 Then
       
            If mstrType = "" Then
                Exit Sub
            End If
            '设置过滤结果分类树表
            Me.tvwClass.Visible = False
            Set objNode = Me.tvwClass.Nodes.Add(, , "_ALL", "所有过滤结果", "close")
            If mstrType = "7" Then
                Set objNode = Me.tvwClass.Nodes.Add("_ALL", tvwChild, "_ALL7", "中草药", "close")
            ElseIf mstrType = "5,6" Then
                Set objNode = Me.tvwClass.Nodes.Add("_ALL", tvwChild, "_ALL5", "西成药", "close")
                Set objNode = Me.tvwClass.Nodes.Add("_ALL", tvwChild, "_ALL6", "中成药", "close")
            ElseIf mstrType = "5" Then
                Set objNode = Me.tvwClass.Nodes.Add("_ALL", tvwChild, "_ALL5", "西成药", "close")
            ElseIf mstrType = "6" Then
                Set objNode = Me.tvwClass.Nodes.Add("_ALL", tvwChild, "_ALL6", "西成药", "close")
            End If
            
            gstrSql = "select Distinct A.ID,A.上级ID,A.编码,A.名称,A.简码,B.类别,Nvl(To_Char(A.撤档时间, 'YYYY-MM-DD'), '3000-01-01') 撤档时间 " & _
                " From 诊疗分类目录 A,诊疗项目目录 B,Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C,Table(Cast(f_Num2List([2]) As zlTools.t_NumList)) D " & _
                " Where A.id=B.分类id And B.类别=C.Column_Value And B.id=D.Column_Value " & IIf(mnuViewList.Checked = False, " And Nvl(To_Char(A.撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' ", "")
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrType, mstrDrugId)
            
            With rsTemp
                Do While Not .EOF
                    Set objNode = Me.tvwClass.Nodes.Add("_ALL" & !类别, tvwChild, "_" & !ID, "[" & !编码 & "]" & !名称, "close")
                    objNode.Sorted = True
                    objNode.Tag = IIf(IsNull(!简码), "", !简码)
                    objNode.ExpandedImage = "expend"
                    If Format(!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                        objNode.ForeColor = mconColor_Stop
                    End If
                    .MoveNext
                Loop
                Me.tvwClass.Visible = True
            End With
            Call setMenu自管药
        Else
            Me.tvwClass.Visible = False
            
            Set objNode = Me.tvwClass.Nodes.Add(, , "_抗菌药", "1-抗菌药", "close")
            objNode.Expanded = True
            Set objNode = Me.tvwClass.Nodes.Add(, , "_原研药", "2-原研药", "close")
            Set objNode = Me.tvwClass.Nodes.Add(, , "_专利药", "3-专利药", "close")
            Set objNode = Me.tvwClass.Nodes.Add(, , "_单独定价", "4-单独定价", "close")
            
            gstrSql = "Select Distinct a.Id, a.上级id, a.编码, a.名称, a.简码, b.类别, Nvl(e.抗生素, 0) 抗生素, Nvl(e.是否原研药, 0) 是否原研药, Nvl(e.是否专利药, 0) 是否专利药," & vbNewLine & _
                      " Nvl(e.是否单独定价, 0) 是否单独定价, Nvl(To_Char(a.撤档时间, 'YYYY-MM-DD'), '3000-01-01') 撤档时间" & vbNewLine & _
                      " From 诊疗分类目录 A, 诊疗项目目录 B, 药品特性 E" & vbNewLine & _
                      " where a.Id = b.分类id And e.药名id = b.Id And Nvl(To_Char(a.撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
            
            '抗菌药
            Set objNode = Me.tvwClass.Nodes.Add("_抗菌药", tvwChild, "_Limit1", "1-非限制使用", "close")
            Set objNode = Me.tvwClass.Nodes.Add("_抗菌药", tvwChild, "_Limit2", "2-限制使用", "close")
            Set objNode = Me.tvwClass.Nodes.Add("_抗菌药", tvwChild, "_Limit3", "3-特殊使用", "close")
            rsTemp.Filter = ""
            rsTemp.Filter = "抗生素<>0"
            With rsTemp
                Do While Not .EOF
                    For i = 1 To Me.tvwClass.Nodes.Count
                        If Me.tvwClass.Nodes(i).Key = "A" & !抗生素 & !ID Then
                            .MoveNext
                            i = 1
                            If .EOF Then Exit For
                        End If
                    Next
                    If Not .EOF Then
                        Set objNode = Me.tvwClass.Nodes.Add("_Limit" & !抗生素, tvwChild, "A" & !抗生素 & !ID, "[" & !编码 & "]" & !名称, "close")
                        objNode.Sorted = True
                        objNode.Tag = IIf(IsNull(!简码), "", !简码)
                        objNode.ExpandedImage = "expend"
                        If Format(!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                            objNode.ForeColor = mconColor_Stop
                        End If
                        .MoveNext
                    End If
                Loop
            End With
            '原研药
            rsTemp.Filter = ""
            rsTemp.Filter = "是否原研药=1"
            With rsTemp
                Do While Not .EOF
                    For i = 1 To Me.tvwClass.Nodes.Count
                        If Me.tvwClass.Nodes(i).Key = "B" & !ID Then
                            .MoveNext
                            i = 1
                            If .EOF Then Exit For
                        End If
                    Next
                    If Not .EOF Then
                        Set objNode = Me.tvwClass.Nodes.Add("_原研药", tvwChild, "B" & !ID, "[" & !编码 & "]" & !名称, "close")
                        objNode.Sorted = True
                        objNode.Tag = IIf(IsNull(!简码), "", !简码)
                        objNode.ExpandedImage = "expend"
                        If Format(!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                            objNode.ForeColor = mconColor_Stop
                        End If
                        .MoveNext
                    End If
                Loop
            End With
            '专利药
            rsTemp.Filter = ""
            rsTemp.Filter = "是否专利药=1"
            With rsTemp
                Do While Not .EOF
                    For i = 1 To Me.tvwClass.Nodes.Count
                        If Me.tvwClass.Nodes(i).Key = "C" & !ID Then
                            .MoveNext
                            i = 1
                            If .EOF Then Exit For
                        End If
                    Next
                    If Not .EOF Then
                        Set objNode = Me.tvwClass.Nodes.Add("_专利药", tvwChild, "C" & !ID, "[" & !编码 & "]" & !名称, "close")
                        objNode.Sorted = True
                        objNode.Tag = IIf(IsNull(!简码), "", !简码)
                        objNode.ExpandedImage = "expend"
                        If Format(!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                            objNode.ForeColor = mconColor_Stop
                        End If
                        .MoveNext
                    End If
                Loop
            End With
            '单独定价
            rsTemp.Filter = ""
            rsTemp.Filter = "是否单独定价=1"
            With rsTemp
                Do While Not .EOF
                    For i = 1 To Me.tvwClass.Nodes.Count
                        If Me.tvwClass.Nodes(i).Key = "D" & !ID Then
                            .MoveNext
                            i = 1
                            If .EOF Then Exit For
                        End If
                    Next
                    If Not .EOF Then
                        Set objNode = Me.tvwClass.Nodes.Add("_单独定价", tvwChild, "D" & !ID, "[" & !编码 & "]" & !名称, "close")
                        objNode.Sorted = True
                        objNode.Tag = IIf(IsNull(!简码), "", !简码)
                        objNode.ExpandedImage = "expend"
                        If Format(!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                            objNode.ForeColor = mconColor_Stop
                        End If
                        .MoveNext
                    End If
                Loop
            End With
            
            Me.tvwClass.Visible = True
            
            Call setMenu自管药
        End If
        Me.stbThis.Panels(2).Text = ""
        If Me.tvwClass.Nodes.Count > 0 Then
            If Val(Me.tvwClass.Tag) = 4 Then
                Me.tvwClass.Nodes("_ALL").Selected = True
            Else
                Me.tvwClass.Nodes("_Limit1").Selected = True
            End If
            Call zlRefRecords
        End If
        Exit Sub
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '填写分类

    gstrSql = "select ID,上级ID,编码,名称,简码,Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') 撤档时间 " & _
            " From 诊疗分类目录" & _
            " Where 类型 = [1] " & IIf(mnuViewList.Checked = False, " And Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' ", "") & _
            " start with 上级ID is null" & _
            " connect by prior ID=上级ID Order By Level, 编码 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, 1 + Val(Me.tvwClass.Tag))
    
    With rsTemp
        Me.tvwClass.Visible = False
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!上级ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !编码 & "]" & !名称, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !上级ID, tvwChild, "_" & !ID, "[" & !编码 & "]" & !名称, "close")
            End If
            objNode.Sorted = True
            objNode.Tag = IIf(IsNull(!简码), "", !简码)
            objNode.ExpandedImage = "expend"
            If Format(!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                objNode.ForeColor = mconColor_Stop
            End If
            .MoveNext
        Loop
        Me.tvwClass.Visible = True
    End With
    If Me.tvwClass.Nodes.Count > 0 Then
        If lngNode <> 0 Then
            Me.tvwClass.Nodes("_" & lngNode).Selected = True
        Else
            Me.tvwClass.Nodes(1).Selected = True
        End If
        Call zlRefRecords
    End If
    Call setMenu自管药
    Exit Sub
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub zlRefRecords(Optional lngItem As Long)
    Dim objListitem As ListItem
    
    '---------------------------------------------
    '填写药品列表
    '---------------------------------------------
    err = 0: On Error GoTo ErrHand
   
    If Val(Me.tvwClass.Tag) <= 2 Then
        gstrSql = "select I.ID,I.编码,I.名称,I.计算单位,T.药品剂型," & _
                "        decode(I.服务对象,1,'门诊',2,'住院',3,'门诊和住院','不直接应用于病人') as 服务对象," & _
                "        decode(T.药品类型,1,'处方药',2,'甲类非处方药',3,'乙类非处方药',4,'非处方药',5,'其它药品',' ') as 药品类型," & _
                "        to_char(nvl(T.处方限量,0)) as 处方限量,decode(T.是否皮试,1,'需要',' ') as 是否皮试," & _
                "        T.毒理分类,T.货源情况,T.价值分类,T.用药梯次," & _
                "        decode(T.是否原料,1,'是',' ') as 是否原料," & _
                "        decode(T.急救药否,1,'是',' ') as 急救药否," & _
                "        decode(I.单独应用,1,'是',' ') as 单独应用," & _
                "        decode(T.是否新药,1,'是',' ') as 是否新药," & _
                "        decode(T.品种医嘱,1,'是',' ') as 品种医嘱," & _
                "        decode(T.是否原研药,1,'是',' ') as 是否原研药," & _
                "        decode(T.是否专利药,1,'是',' ') as 是否专利药," & _
                "        decode(T.是否单独定价,1,'是',' ') as 是否单独定价," & _
                "        decode(nvl(t.抗生素,0),0,'',1,'非限制使用',2,'限制使用','特殊使用') as 抗菌药物," & _
                "        decode(T.是否辅助用药,1,'是',' ') as 是否辅助用药," & _
                "        nvl(I.撤档时间,to_date('3000-01-01','YYYY-MM-DD')) as 撤档时间,Nvl(I.适用性别,0) As 适用性别 " & _
                " from 诊疗项目目录 I,药品特性 T" & _
                " where I.ID=T.药名ID and "
        If mnuViewShowAll.Checked = False Then
            gstrSql = gstrSql & " I.分类ID=[1] "
        Else
            gstrSql = gstrSql & " I.分类ID IN " & _
                " (Select ID From 诊疗分类目录 Where 类型 In (1,2,3) " & _
                "  Start With ID=[1] Connect By Prior ID=上级ID)"
        End If
        If Me.mnuViewStoped.Checked = False Then
            gstrSql = gstrSql & " and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
        End If
        If mbln自管药 = True Then
            gstrSql = gstrSql & " and t.临床自管药=1"
        Else
            gstrSql = gstrSql & " and t.临床自管药 is null"
        End If
        gstrSql = gstrSql & " order by I.编码"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Me.tvwClass.SelectedItem.Key, 2)))
        
        With rsTemp
            Me.lvwItems.ListItems.Clear
            Do While Not .EOF
                Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
                If Format(!撤档时间, "YYYY-MM-DD") = "3000-01-01" Then
                    If Val(Me.tvwClass.Tag) = 2 Then
                        objItem.Icon = "草药U": objItem.SmallIcon = "草药U"
                    Else
                        objItem.Icon = "成药U": objItem.SmallIcon = "成药U"
                    End If
                Else
                    If Val(Me.tvwClass.Tag) = 2 Then
                        objItem.Icon = "草药S": objItem.SmallIcon = "草药S"
                    Else
                        objItem.Icon = "成药S": objItem.SmallIcon = "成药S"
                    End If
                End If
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_编码").Index - 1) = !编码
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_剂量单位").Index - 1) = IIf(IsNull(!计算单位), "", !计算单位)
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_剂型").Index - 1) = IIf(IsNull(!药品剂型), "", !药品剂型)
                'objItem.SubItems(Me.lvwItems.ColumnHeaders("_服务对象").Index - 1) = !服务对象
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_处方类型").Index - 1) = !药品类型
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_处方限量").Index - 1) = !处方限量
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_过敏试验").Index - 1) = !是否皮试
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_毒理").Index - 1) = !毒理分类
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_货源").Index - 1) = !货源情况
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_价值").Index - 1) = !价值分类
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_梯次").Index - 1) = !用药梯次
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_原料药").Index - 1) = !是否原料
                If Val(Me.tvwClass.Tag) = 2 Then
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_单味使用").Index - 1) = !单独应用
                Else
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_急救药").Index - 1) = !急救药否
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_新药").Index - 1) = !是否新药
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_原研药").Index - 1) = !是否原研药
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_专利药").Index - 1) = !是否专利药
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_单独定价").Index - 1) = !是否单独定价
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_抗菌药物").Index - 1) = zlStr.Nvl(!抗菌药物, "")
                End If
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_按药品下长期医嘱").Index - 1) = !品种医嘱
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_辅助用药").Index - 1) = !是否辅助用药
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_适用性别").Index - 1) = IIf(!适用性别 = 1, "男性", IIf(!适用性别 = 2, "女性", "无性别区分"))
                If !ID = lngItem Then
                    objItem.Selected = True
                End If
                If Format(!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                    objItem.ForeColor = mconColor_Stop
                    For intCount = 1 To Me.lvwItems.ColumnHeaders.Count - 1
                        objItem.ListSubItems(intCount).ForeColor = mconColor_Stop
                    Next
                End If
                .MoveNext
            Loop
        End With
    Else        '单独处理过滤结果
        Dim strType As String
        
        strType = Mid(Me.tvwClass.SelectedItem.Key, 2)      '根据所选节点来生成查询条件
        
        If Not IsNumeric(strType) Then                      '如果不是数字就是汇总条件
            If strType = "ALL" Then
                strType = " And I.类别 in('5','6','7') "
            ElseIf strType = "ALL5" Then
                 strType = " And I.类别='5' "
            ElseIf strType = "ALL6" Then
                strType = " And I.类别='6' "
            ElseIf strType = "ALL7" Then
                strType = " And I.类别='7' "
            Else
                Select Case strType
                    Case "抗菌药"
                        strType = " And nvl(T.抗生素,0)<>0 "
                    Case "原研药"
                        strType = " And nvl(T.是否原研药,0)<>0 "
                    Case "专利药"
                        strType = " And nvl(T.是否专利药,0)<>0 "
                    Case "单独定价"
                        strType = " And nvl(T.是否单独定价,0)<>0 "
                    Case "Limit1", "Limit2", "Limit3"
                        strType = " And nvl(T.抗生素,0)=" & Mid(strType, 6) & " "
                End Select
            End If
        Else                                                '是数字就是分类ID条件
            If Val(Me.tvwClass.Tag) = 4 Then
                strType = " And I.分类id=" & strType & " "
            Else
                Select Case Mid(Me.tvwClass.SelectedItem.Key, 1, 1)
                    Case "A"
                        strType = " And nvl(T.抗生素,0)=" & Mid(strType, 1, 1) & " And I.分类id=" & Mid(strType, 2) & " "
                    Case "B"
                        strType = " And nvl(T.是否原研药,0)<>0 And I.分类id=" & Mid(strType, 1) & " "
                    Case "C"
                        strType = " And nvl(T.是否专利药,0)<>0 And I.分类id=" & Mid(strType, 1) & " "
                    Case "D"
                        strType = " And nvl(T.是否单独定价,0)<>0 And I.分类id=" & Mid(strType, 1) & " "
                End Select
            End If
        End If
        
        Me.lvwItems.Visible = False
        
        gstrSql = "select I.ID,I.编码,I.名称,I.计算单位,T.药品剂型," & _
                "        decode(I.服务对象,1,'门诊',2,'住院',3,'门诊和住院','不直接应用于病人') as 服务对象," & _
                "        decode(T.药品类型,1,'处方药',2,'甲类非处方药',3,'乙类非处方药',4,'非处方药',5,'其它药品',' ') as 药品类型," & _
                "        to_char(nvl(T.处方限量,0)) as 处方限量,decode(T.是否皮试,1,'需要',' ') as 是否皮试," & _
                "        T.毒理分类,T.货源情况,T.价值分类,T.用药梯次," & _
                "        decode(T.是否原料,1,'是',' ') as 是否原料," & _
                "        decode(T.急救药否,1,'是',' ') as 急救药否," & _
                "        decode(I.单独应用,1,'是',' ') as 单独应用," & _
                "        decode(T.是否新药,1,'是',' ') as 是否新药," & _
                "        decode(T.品种医嘱,1,'是',' ') as 品种医嘱," & _
                "        decode(T.是否原研药,1,'是',' ') as 是否原研药," & _
                "        decode(T.是否专利药,1,'是',' ') as 是否专利药," & _
                "        decode(T.是否单独定价,1,'是',' ') as 是否单独定价," & _
                "        decode(I.类别,'5','1','6','2','3') as 类型," & _
                "        I.分类id," & _
                "        decode(T.是否辅助用药,1,'是',' ') as 是否辅助用药," & _
                "        nvl(I.撤档时间,to_date('3000-01-01','YYYY-MM-DD')) as 撤档时间,Nvl(I.适用性别,0) As 适用性别, " & _
                "        decode(nvl(t.抗生素,0),0,'',1,'非限制使用',2,'限制使用','特殊使用') as 抗菌药物 " & _
                " from 诊疗项目目录 I,药品特性 T" & _
                " where I.ID=T.药名ID " & IIf(Val(Me.tvwClass.Tag) = 4, IIf(mstrDrugId <> "", "And I.ID IN ( " & mstrDrugId & ") ", ""), "") & strType
        
        If mbln自管药 = True Then
            gstrSql = gstrSql & " and t.临床自管药=1"
        Else
            gstrSql = gstrSql & " and t.临床自管药 is null"
        End If
    
        If Me.mnuViewStoped.Checked = False Then
            gstrSql = gstrSql & " and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
        End If
        gstrSql = gstrSql & " order by I.编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "zlRefRecords")

        Me.lvwItems.ListItems.Clear
        With rsTemp
            Do While Not .EOF
                Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
                If Format(!撤档时间, "YYYY-MM-DD") = "3000-01-01" Then
                    If mstrType = "7" Then
                        objItem.Icon = "草药U": objItem.SmallIcon = "草药U"
                    Else
                        objItem.Icon = "成药U": objItem.SmallIcon = "成药U"
                    End If
                Else
                    If mstrType = "7" Then
                        objItem.Icon = "草药S": objItem.SmallIcon = "草药S"
                    Else
                        objItem.Icon = "成药S": objItem.SmallIcon = "成药S"
                    End If
                End If
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_编码").Index - 1) = !编码
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_剂量单位").Index - 1) = IIf(IsNull(!计算单位), "", !计算单位)
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_剂型").Index - 1) = IIf(IsNull(!药品剂型), "", !药品剂型)
                'objItem.SubItems(Me.lvwItems.ColumnHeaders("_服务对象").Index - 1) = !服务对象
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_处方类型").Index - 1) = !药品类型
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_处方限量").Index - 1) = !处方限量
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_过敏试验").Index - 1) = !是否皮试
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_毒理").Index - 1) = !毒理分类
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_货源").Index - 1) = !货源情况
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_价值").Index - 1) = !价值分类
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_梯次").Index - 1) = !用药梯次
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_原料药").Index - 1) = !是否原料
                If mstrType = "7" Then
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_单味使用").Index - 1) = !单独应用
                Else
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_急救药").Index - 1) = !急救药否
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_新药").Index - 1) = !是否新药
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_原研药").Index - 1) = !是否原研药
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_专利药").Index - 1) = !是否专利药
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_单独定价").Index - 1) = !是否单独定价
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_抗菌药物").Index - 1) = zlStr.Nvl(!抗菌药物, "")
                End If
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_按药品下长期医嘱").Index - 1) = !品种医嘱
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_适用性别").Index - 1) = IIf(!适用性别 = 1, "男性", IIf(!适用性别 = 2, "女性", "无性别区分"))
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_类型").Index - 1) = !类型
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_分类id").Index - 1) = !分类id
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_辅助用药").Index - 1) = !是否辅助用药
                If !ID = lngItem Then
                    objItem.Selected = True
                End If
                If Format(!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                    objItem.ForeColor = mconColor_Stop
                    For intCount = 1 To Me.lvwItems.ColumnHeaders.Count - 1
                        objItem.ListSubItems(intCount).ForeColor = mconColor_Stop
                    Next
                End If
                .MoveNext
            Loop
        End With
        Me.lvwItems.Visible = True
    End If
    
    For Each objItem In lvwItems.ListItems
        If objItem.Key = mStrItem Then
            lvwItems.ListItems(objItem.Key).Selected = True
            Exit For
        End If
    Next
    
    If Me.lvwItems.ListItems.Count > 0 Then
        If Me.lvwItems.SelectedItem Is Nothing Then Me.lvwItems.ListItems(1).Selected = True
        Call lvwItems_ItemClick(lvwItems.SelectedItem)
        err = 0: On Error Resume Next
        DoEvents: Me.lvwItems.SelectedItem.EnsureVisible
        Me.stbThis.Panels(2).Text = "该分类共有" & Me.lvwItems.ListItems.Count & "种药品"
    Else
        Me.lvwSpecs.ListItems.Clear
        With Me.hgdPrice
            .Redraw = False
            .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
            For intCol = 0 To .Cols - 1
                .TextMatrix(.FixedRows, intCol) = ""
            Next
            .Redraw = True
        End With
        For intCount = Me.lblComment.LBound To Me.lblComment.UBound
            Me.lblComment(intCount).Caption = ""
        Next
        
        If fraComment(0).Caption = "" Then
            lvwSpecs.Height = tabContent.Height - 450
        End If
        Me.stbThis.Panels(2).Text = ""
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub zlPopupEditMenu(bytEditKind As Byte, blnStopUse As Boolean)
    '-------------------------------------------------
    '功能:弹出编辑菜单
    '入参:  bytEditKind:1-品种编辑 2-规格编辑
    '       blnStopUse:是否包括停用和启用功能
    '-------------------------------------------------
    Dim objItem As ListItem
    Dim StrClass As String
    
    On Error GoTo RESHOW
    Me.mnuEditItemAdd.Tag = Me.mnuEditItemAdd.Visible
    Me.mnuEditItemMod.Tag = Me.mnuEditItemMod.Visible
    Me.mnuEditItemDel.Tag = Me.mnuEditItemDel.Visible
    Me.mnuEditItemTabu.Tag = Me.mnuEditItemTabu.Visible
    Me.mnuEditItemPart.Tag = Me.mnuEditItemPart.Visible
    Me.mnuPriceChargeSet.Tag = Me.mnuPriceChargeSet.Visible
    Me.mnuEditSpt1.Tag = Me.mnuEditSpt1.Visible
    Me.mnuEditSpecAdd.Tag = Me.mnuEditSpecAdd.Visible
    Me.mnuEditSpecMod.Tag = Me.mnuEditSpecMod.Visible
    Me.mnuEditSpecDel.Tag = Me.mnuEditSpecDel.Visible
    Me.mnuEditSpecLimit.Tag = Me.mnuEditSpecLimit.Visible
    Me.mnuEditSendType.Tag = Me.mnuEditSendType.Visible
    Me.mnuEditSpecProtocol.Tag = Me.mnuEditSpecProtocol.Visible
    Me.mnuEditSpecSelf.Tag = Me.mnuEditSpecSelf.Visible
    Me.mnuEditSpt2.Tag = Me.mnuEditSpt2.Visible
    Me.mnuEditStart.Tag = Me.mnuEditStart.Visible
    Me.mnuEditStop.Tag = Me.mnuEditStop.Visible
    Me.mnuEditSptPacker.Tag = Me.mnuEditSptPacker.Visible
    Me.mnuUploadDrugInfo.Tag = Me.mnuUploadDrugInfo.Visible
    
    Me.mnuPriceChargeSet1.Visible = Me.mnuPriceChargeSet.Visible
    Me.mnuEditSpt3.Visible = Me.mnuPriceChargeSet.Visible
    
    Me.mnuPriceChargeSet1.Enabled = Me.mnuPriceChargeSet.Enabled
    
    Select Case bytEditKind
    Case 1  '品种
        If InStr(1, mstrPrivs, ";对应处方;") = 0 Then
            mnuEditItemBill.Visible = False
        Else
            mnuEditItemBill.Visible = True
        End If
        If InStr(1, mstrPrivs, ";配伍禁忌关系;") = 0 Then
            mnuEditItemTabu.Visible = False
        Else
            mnuEditItemTabu.Visible = True
        End If
        With lvwItems
            If tvwClass.Tag >= 3 Then
                Set objItem = .SelectedItem
                If objItem.SubItems(.ColumnHeaders("_类型").Index - 1) = "1" Then
                    StrClass = "5"
                ElseIf objItem.SubItems(.ColumnHeaders("_类型").Index - 1) = "2" Then
                    StrClass = "6"
                ElseIf objItem.SubItems(.ColumnHeaders("_类型").Index - 1) = "3" Then
                    StrClass = "7"
                End If
            Else
                If tvwClass.Tag = 0 Then
                    StrClass = "5"
                ElseIf tvwClass.Tag = 1 Then
                    StrClass = "6"
                ElseIf tvwClass.Tag = 2 Then
                    StrClass = "7"
                End If
            End If
            
            Select Case StrClass
            Case "5" '西成药
                If InStr(1, mstrPrivs, ";管理西成药品种;") = 0 Then
                    mnuEditItemAdd.Visible = False
                    mnuEditItemMod.Visible = False
                    mnuEditItemDel.Visible = False
                    mnuEditStart.Visible = False
                    mnuEditStop.Visible = False
                Else
                    mnuEditItemAdd.Visible = True
                    mnuEditItemMod.Visible = True
                    mnuEditItemDel.Visible = True
                    
                    If InStr(1, mstrPrivs, ";药品启用;") = 0 Then
                        mnuEditStart.Visible = False
                    Else
                        mnuEditStart.Visible = True
                    End If
                    If InStr(1, mstrPrivs, ";药品停用;") = 0 Then
                        mnuEditStop.Visible = False
                    Else
                        mnuEditStop.Visible = True
                    End If
                End If
                If InStr(1, mstrPrivs, ";用法用量;") = 0 Then
                    mnuEditItemUsage.Visible = False
                Else
                    mnuEditItemUsage.Visible = True
                End If
                mnuEditItemTabu.Enabled = True
                mnuEditItemUsage.Enabled = True
                mnuEditItemBill.Enabled = True
            Case "6" '中成药
                If InStr(1, mstrPrivs, ";管理中成药品种;") = 0 Then
                    mnuEditItemAdd.Visible = False
                    mnuEditItemMod.Visible = False
                    mnuEditItemDel.Visible = False
                    mnuEditStart.Visible = False
                    mnuEditStop.Visible = False
                Else
                    mnuEditItemAdd.Visible = True
                    mnuEditItemMod.Visible = True
                    mnuEditItemDel.Visible = True
                    
                    If InStr(1, mstrPrivs, ";药品启用;") = 0 Then
                        mnuEditStart.Visible = False
                    Else
                        mnuEditStart.Visible = True
                    End If
                    If InStr(1, mstrPrivs, ";药品停用;") = 0 Then
                        mnuEditStop.Visible = False
                    Else
                        mnuEditStop.Visible = True
                    End If
                End If
                If InStr(1, mstrPrivs, ";用法用量;") = 0 Then
                    mnuEditItemUsage.Visible = False
                Else
                    mnuEditItemUsage.Visible = True
                End If
                mnuEditItemTabu.Enabled = True
                mnuEditItemUsage.Enabled = True
                mnuEditItemBill.Enabled = True
            Case "7"   '中草药
                If InStr(1, mstrPrivs, ";管理中草药品种;") = 0 Then
                    mnuEditItemAdd.Visible = False
                    mnuEditItemMod.Visible = False
                    mnuEditItemDel.Visible = False
                    mnuEditStart.Visible = False
                    mnuEditStop.Visible = False
                Else
                    mnuEditItemAdd.Visible = True
                    mnuEditItemMod.Visible = True
                    mnuEditItemDel.Visible = True
                    
                    If InStr(1, mstrPrivs, ";药品启用;") = 0 Then
                        mnuEditStart.Visible = False
                    Else
                        mnuEditStart.Visible = True
                    End If
                    If InStr(1, mstrPrivs, ";药品停用;") = 0 Then
                        mnuEditStop.Visible = False
                    Else
                        mnuEditStop.Visible = True
                    End If
                End If
                mnuEditItemUsage.Visible = False '用法用量
                mnuEditItemTabu.Enabled = True
                mnuEditItemBill.Enabled = True
            End Select
        End With
        
        If tvwClass.Nodes.Count > 0 Then    '当有分类时
            mnuEditItemAdd.Enabled = True   '新增品种
        End If
        mnuEditItemDel.Enabled = True   '删除品种
        If lvwItems.ListItems.Count > 0 Then '有品种时
            If lvwItems.SelectedItem.Icon Like "*U" = True Then  '未停用品种
                mnuEditItemMod.Enabled = True   '修改品种
                mnuEditStart.Enabled = False     '启用
                mnuEditStop.Enabled = True     '停用
            ElseIf lvwItems.SelectedItem.Icon Like "*S" = True Then   '已停用
                mnuEditItemMod.Enabled = False   '修改品种
                mnuEditStart.Enabled = True     '启用
                mnuEditStop.Enabled = False     '停用
            End If
        Else
            mnuEditItemAdd.Enabled = True
            mnuEditItemMod.Enabled = False
            mnuEditItemDel.Enabled = False
        End If
        mnuEditSpecAdd.Visible = False
        mnuEditSpecMod.Visible = False
        mnuEditSpecDel.Visible = False
        mnuEditSpt1.Visible = False
        mnuEditSpt7.Visible = False
        mnuEditItemPart.Visible = False
        mnuEditSpecLimit.Visible = False
        mnuEditSpecProtocol.Visible = False
        mnuEditSpecSelf.Visible = False
        mnuEditSpecUnit.Visible = False
        mnuEditManFac.Visible = False
        mnuEditSendType.Visible = False
        mnuPriceChargeSet1.Visible = False   '费别设置
        mnuEditSpt4.Visible = False
        mnuEditSpt3.Visible = mnuEditStop.Visible
    Case 2  '规格
        With lvwItems
            If tvwClass.Tag >= 3 Then
                Set objItem = .SelectedItem
                If objItem.SubItems(.ColumnHeaders("_类型").Index - 1) = "1" Then
                    StrClass = "5"
                ElseIf objItem.SubItems(.ColumnHeaders("_类型").Index - 1) = "2" Then
                    StrClass = "6"
                ElseIf objItem.SubItems(.ColumnHeaders("_类型").Index - 1) = "3" Then
                    StrClass = "7"
                End If
            Else
                If tvwClass.Tag = 0 Then
                    StrClass = "5"
                ElseIf tvwClass.Tag = 1 Then
                    StrClass = "6"
                ElseIf tvwClass.Tag = 2 Then
                    StrClass = "7"
                End If
            End If
            
            Select Case StrClass
            Case "5" '西成药
                If InStr(1, mstrPrivs, ";管理西成药规格;") = 0 Then
                    mnuEditSpecAdd.Visible = False
                    mnuEditSpecMod.Visible = False
                    mnuEditSpecDel.Visible = False
                    mnuEditStart.Visible = False
                    mnuEditStop.Visible = False
                Else
                    mnuEditSpecAdd.Visible = True
                    mnuEditSpecMod.Visible = True
                    mnuEditSpecDel.Visible = True
                    
                    If InStr(1, mstrPrivs, ";药品启用;") = 0 Then
                        mnuEditStart.Visible = False
                    Else
                        mnuEditStart.Visible = True
                    End If
                    If InStr(1, mstrPrivs, ";药品停用;") = 0 Then
                        mnuEditStop.Visible = False
                    Else
                        mnuEditStop.Visible = True
                    End If
                    
                    If InStr(1, mstrPrivs, ";用法用量;") = 0 Then
                        mnuEditItemUsage.Visible = False
                    Else
                        mnuEditItemUsage.Visible = True
                        mnuEditItemUsage.Enabled = True
                    End If
                End If
            Case "6" '中成药
                If InStr(1, mstrPrivs, ";管理中成药规格;") = 0 Then
                    mnuEditSpecAdd.Visible = False
                    mnuEditSpecMod.Visible = False
                    mnuEditSpecDel.Visible = False
                    mnuEditStart.Visible = False
                    mnuEditStop.Visible = False
                Else
                    mnuEditSpecAdd.Visible = True
                    mnuEditSpecMod.Visible = True
                    mnuEditSpecDel.Visible = True
                    
                    If InStr(1, mstrPrivs, ";药品启用;") = 0 Then
                        mnuEditStart.Visible = False
                    Else
                        mnuEditStart.Visible = True
                    End If
                    If InStr(1, mstrPrivs, ";药品停用;") = 0 Then
                        mnuEditStop.Visible = False
                    Else
                        mnuEditStop.Visible = True
                    End If
                    
                    If InStr(1, mstrPrivs, ";用法用量;") = 0 Then
                        mnuEditItemUsage.Visible = False
                    Else
                        mnuEditItemUsage.Visible = True
                        mnuEditItemUsage.Enabled = True
                    End If
                End If
            Case "7"   '中草药
                If InStr(1, mstrPrivs, ";管理中草药规格;") = 0 Then
                    mnuEditSpecAdd.Visible = False
                    mnuEditSpecMod.Visible = False
                    mnuEditSpecDel.Visible = False
                    mnuEditStart.Visible = False
                    mnuEditStop.Visible = False
                Else
                    mnuEditSpecAdd.Visible = True
                    mnuEditSpecMod.Visible = True
                    mnuEditSpecDel.Visible = True
                    
                    If InStr(1, mstrPrivs, ";药品启用;") = 0 Then
                        mnuEditStart.Visible = False
                    Else
                        mnuEditStart.Visible = True
                    End If
                    If InStr(1, mstrPrivs, ";药品停用;") = 0 Then
                        mnuEditStop.Visible = False
                    Else
                        mnuEditStop.Visible = True
                    End If
                End If
            End Select
            If InStr(1, mstrPrivs, ";费别设置;") = 0 Then
                mnuPriceChargeSet1.Visible = False
                mnuEditSpt4.Visible = False
            Else
                mnuPriceChargeSet1.Visible = True
                mnuEditSpt4.Visible = True
            End If
        End With
        
        If lvwItems.ListItems.Count > 0 And lvwItems.SelectedItem.Icon Like "*U" = True Then '有品种,未停用
            mnuEditSpecAdd.Enabled = True
        End If
        mnuEditSpecDel.Enabled = True   '删除规格
        If lvwSpecs.ListItems.Count > 0 Then '有规格时
            If lvwSpecs.SelectedItem.Icon Like "*U" = True Then '未停用
                mnuEditSpecMod.Enabled = True   '修改规格
                mnuEditStart.Enabled = False     '启用
                mnuEditStop.Enabled = True     '停用
            ElseIf lvwSpecs.SelectedItem.Icon Like "*S" = True Then '已停用
                mnuEditSpecMod.Enabled = False   '修改规格
                mnuEditStart.Enabled = True     '启用
                mnuEditStop.Enabled = False     '停用
            End If
        Else
            mnuEditSpecDel.Enabled = False
            mnuEditSpecMod.Enabled = False
            mnuEditStart.Enabled = False     '启用
            mnuEditStop.Enabled = False     '停用
        End If
        mnuEditItemAdd.Visible = False
        mnuEditItemMod.Visible = False
        mnuEditItemDel.Visible = False
        mnuEditItemTabu.Visible = False '配伍禁忌
'        mnuEditItemUsage.Visible = False '用法用量
        mnuEditItemBill.Visible = False '对应处方
        mnuEditSpt1.Visible = False
        mnuEditSpt7.Visible = mnuEditSpecAdd.Visible
        mnuEditItemPart.Visible = True
        mnuEditSpecLimit.Visible = True
        mnuEditSpecProtocol.Visible = True
        mnuEditSpecSelf.Visible = True
        mnuEditSpecUnit.Visible = True
        mnuEditManFac.Visible = True
        mnuEditSendType.Visible = True
        mnuEditSpt3.Visible = mnuEditStop.Visible
    End Select
    
    Call setMenu自管药
    Call PopupMenu(Me.mnuEdit, 2)
    
RESHOW:
    Me.mnuEditItemAdd.Visible = Me.mnuEditItemAdd.Tag
    Me.mnuEditItemMod.Visible = Me.mnuEditItemMod.Tag
    Me.mnuEditItemDel.Visible = Me.mnuEditItemDel.Tag
    Me.mnuEditItemTabu.Visible = Me.mnuEditItemTabu.Tag
    Me.mnuEditItemPart.Visible = Me.mnuEditItemPart.Tag
    Me.mnuPriceChargeSet.Visible = Me.mnuPriceChargeSet.Tag
    Me.mnuEditSpt1.Visible = Me.mnuEditSpt1.Tag
    Me.mnuEditSpecAdd.Visible = Me.mnuEditSpecAdd.Tag
    Me.mnuEditSpecMod.Visible = Me.mnuEditSpecMod.Tag
    Me.mnuEditSpecDel.Visible = Me.mnuEditSpecDel.Tag
    Me.mnuEditSpecLimit.Visible = Me.mnuEditSpecLimit.Tag
    Me.mnuEditSendType.Visible = Me.mnuEditSendType.Tag
    Me.mnuEditSpecProtocol.Visible = Me.mnuEditSpecProtocol.Tag
    Me.mnuEditSpecSelf.Visible = Me.mnuEditSpecSelf.Tag
    Me.mnuEditSpt2.Visible = Me.mnuEditSpt2.Tag
    Me.mnuEditStart.Visible = Me.mnuEditStart.Tag
    Me.mnuEditStop.Visible = Me.mnuEditStop.Tag
    Me.mnuEditSptPacker.Visible = Me.mnuEditSptPacker.Tag
    Me.mnuUploadDrugInfo.Visible = Me.mnuUploadDrugInfo.Tag
    Me.mnuEditSpecUnit.Visible = True

    Call setMenu自管药
End Sub

Private Sub setMenu自管药()
    '功能：自管药菜单控制
    If mbln自管药 = True Then   '自管药菜单控制
        mnuFilePara.Visible = False
        mnuFileSpt2.Visible = False
        mnuClass.Visible = False
        mnuEditItemPart.Visible = False
        mnuEditSpecLimit.Visible = False
        mnuEditSpecProtocol.Visible = False
        mnuEditSpecSelf.Visible = False
        mnuEditSpecUnit.Visible = False
        mnuEditManFac.Visible = False
        mnuEditSendType.Visible = False
        mnuEditSpt5.Visible = False
        mnuEditRate.Visible = False
        mnuEditSpt2.Visible = False
        mnuEditVariBatch.Visible = False
        mnuEditSpecBatch.Visible = False
        mnuEditExcel.Visible = False
        mnuEditSpt4.Visible = False
        mnuPriceChargeSet1.Visible = False
        mnuPriceSpt1.Visible = False
        mnuEditSptPacker.Visible = False
        mnuUploadDrugInfo.Visible = False
        mnuPrice.Visible = False
        mnuEditSpt3.Visible = False
        mnuViewPrices.Visible = False
        tlbThis.Buttons("Limit").Visible = False
        tlbThis.Buttons("Class").Visible = False
        tlbThis.Buttons(10).Visible = False
        mnuEditSpt6.Visible = True
    End If
End Sub
Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '功能:记录表打印
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrintLvw
    err = 0: On Error Resume Next
    Set objPrint.Body.objData = Me.lvwItems
    Select Case Val(Me.tvwClass.Tag)
    Case 0
        objPrint.Title.Text = "西成药品种清单"
    Case 1
        objPrint.Title.Text = "中成药品种清单"
    Case 2
        objPrint.Title.Text = "中草药清单"
    End Select
    objPrint.UnderAppItems.Add "分类：" & Me.tvwClass.SelectedItem.Text
    objPrint.BelowAppItems.Add "打印时间：" & Now
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrViewLvw objPrint, bytMode
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

Public Sub zlLocateItem(lng分类id As Long, lng药名id As Long, lng药品ID As Long)
    Dim lstItem As ListItem, lstSpec As ListItem, tvwNode As Node
    '---------------------------------------------
    '定位到指定的诊断参考项目，在查找时使用
    '---------------------------------------------
    On Error GoTo ErrHand
    Set tvwNode = tvwClass.SelectedItem
    Set lstItem = lvwItems.SelectedItem
    Set lstSpec = lvwSpecs.SelectedItem
    
'    If lstItem Is Nothing Then
'        Exit Sub
'    End If
    '选择分类
    Set Me.tvwClass.SelectedItem = Me.tvwClass.Nodes("_" & lng分类id)
    Me.tvwClass.Nodes("_" & lng分类id).Selected = True
    Me.tvwClass.SelectedItem.EnsureVisible
    Call zlRefRecords
    If lvwItems.ListItems.Count <> 0 Then '如果分类下面没有品种时，没有必要在定位
        '选择品种
        Set Me.lvwItems.SelectedItem = Me.lvwItems.ListItems("_" & lng药名id)
        Me.lvwItems.SelectedItem.EnsureVisible
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
        '选择药品
        If lng药品ID <> 0 Then
            Set Me.lvwSpecs.SelectedItem = Me.lvwSpecs.ListItems("_" & lng药品ID)
            If err <> 0 Then
                Set Me.lvwSpecs.SelectedItem = Me.lvwSpecs.ListItems(1)
            End If
            Me.lvwSpecs.SelectedItem.EnsureVisible
        End If
    End If
    Exit Sub
ErrHand:
    Set tvwClass.SelectedItem = tvwNode
    Call zlRefRecords
    Set Me.lvwItems.SelectedItem = Me.lvwItems.ListItems(lstItem.Key)
    Me.lvwItems.SelectedItem.EnsureVisible
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    Set Me.lvwSpecs.SelectedItem = Me.lvwSpecs.ListItems(lstSpec.Key)
    Me.lvwSpecs.SelectedItem.EnsureVisible
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub


Public Sub ZlRefBut(ByVal intType As Integer)
    If intType = 3 Then
        cmdKind_Click (3)
    End If
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim strTemp As String
    Dim strTag As String
    
    On Error GoTo ErrHandle
    
    If KeyAscii = vbKeyReturn Then
        zlControl.TxtSelAll txtFind
        If txtFind.Text = "" Then Exit Sub
        If mstrFindValue <> txtFind.Text And txtFind.Text <> "" Then
            mstrFindValue = txtFind.Text
            Set mrsFind = Nothing
            
            strTemp = " And (I.撤档时间 Is NULL Or to_Char(I.撤档时间,'yyyy-MM-dd')='3000-01-01')"
            
            If mbln自管药 = False Then
                gstrSql = "SELECT DISTINCT I.分类ID,I.ID AS 药名ID,0 AS 药品ID" & _
                    " FROM 诊疗项目目录 I,诊疗项目别名 N" & _
                    " WHERE I.ID=N.诊疗项目ID " & _
                    " AND I.类别=[1] " & _
                    " AND (I.编码 LIKE [2] " & _
                    "     OR N.名称 LIKE [2] " & _
                    "     OR N.简码 LIKE [2])"
            Else
                gstrSql = "Select Distinct i.分类id, i.Id As 药名id, 0 As 药品id" & vbNewLine & _
                    "From 诊疗项目目录 I, 诊疗项目别名 N, 药品特性 A" & vbNewLine & _
                    "Where i.Id = n.诊疗项目id And i.Id = a.药名id And a.临床自管药 = 1 And i.类别 = [1] And" & vbNewLine & _
                    "      (i.编码 Like [2] Or n.名称 Like [2] Or n.简码 Like [2])"
            End If
            
            If tvwClass.Tag = "0" Then
                strTag = "5"
            ElseIf tvwClass.Tag = "1" Then
                strTag = "6"
            ElseIf tvwClass.Tag = "2" Then
                strTag = "7"
            End If
            If mnuViewStoped.Checked = False Then
                gstrSql = gstrSql & strTemp
            End If
            Set mrsFind = zlDatabase.OpenSQLRecord(gstrSql, "药品查询", strTag, gstrMatch & UCase(txtFind.Text) & "%")
            If mrsFind.RecordCount > 0 Then
                Call zlLocateItem(mrsFind!分类id, mrsFind!药名ID, mrsFind!药品id)
            End If
        Else
            If Not mrsFind.EOF Then
                mrsFind.MoveNext
                If Not mrsFind.EOF Then
                    Call zlLocateItem(mrsFind!分类id, mrsFind!药名ID, mrsFind!药品id)
                Else
                    MsgBox "已查询到最后一条记录！", vbInformation, gstrSysName
                    mrsFind.MoveFirst
                    Call zlLocateItem(mrsFind!分类id, mrsFind!药名ID, mrsFind!药品id)
                End If
            ElseIf mrsFind.RecordCount <> 0 And mrsFind.EOF Then
                mrsFind.MoveFirst
                Call zlLocateItem(mrsFind!分类id, mrsFind!药名ID, mrsFind!药品id)
            End If
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub




