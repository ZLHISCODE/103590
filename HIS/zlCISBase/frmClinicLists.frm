VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmClinicLists 
   BackColor       =   &H8000000C&
   Caption         =   "诊疗项目管理"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12450
   Icon            =   "frmClinicLists.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   12450
   ShowInTaskbar   =   0   'False
   Begin ComCtl3.CoolBar clbThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12450
      _ExtentX        =   21960
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   12450
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "tlbThis"
      MinHeight1      =   720
      Width1          =   10005
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Caption2        =   "查找"
      Child2          =   "txtFind"
      MinHeight2      =   300
      Width2          =   1080
      Key2            =   "find"
      NewRow2         =   0   'False
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   10620
         TabIndex        =   41
         Top             =   240
         Width           =   1740
      End
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   720
         Left            =   165
         TabIndex        =   10
         Top             =   30
         Width           =   9810
         _ExtentX        =   17304
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
            NumButtons      =   16
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
               Key             =   "Split1"
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
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加"
               Key             =   "Add"
               Description     =   "增加"
               Object.ToolTipText     =   "增加新的项目"
               Object.Tag             =   "增加"
               ImageIndex      =   4
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "add"
                     Text            =   "增加"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "addcopy"
                     Text            =   "复制增加"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modify"
               Description     =   "修改"
               Object.ToolTipText     =   "修改当前项目"
               Object.Tag             =   "修改"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Delete"
               Description     =   "删除"
               Object.ToolTipText     =   "删除当前项目"
               Object.Tag             =   "删除"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split3"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "启用"
               Key             =   "Start"
               Description     =   "启用"
               Object.ToolTipText     =   "启用指定的停用项目"
               Object.Tag             =   "启用"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "停用"
               Key             =   "Stop"
               Description     =   "停用"
               Object.ToolTipText     =   "停用指定的在用项目"
               Object.Tag             =   "停用"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split4"
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
               Key             =   "Split5"
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picHBar 
      Appearance      =   0  'Flat
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
      Left            =   2625
      MousePointer    =   7  'Size N S
      ScaleHeight     =   30
      ScaleWidth      =   6075
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5910
      Width           =   6075
   End
   Begin VB.PictureBox picVBar 
      Appearance      =   0  'Flat
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
      Height          =   6270
      Left            =   0
      ScaleHeight     =   6210
      ScaleWidth      =   2340
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   2400
      Begin VB.CommandButton cmdKind 
         Caption         =   "成套方案(&2)"
         Height          =   350
         Index           =   2
         Left            =   0
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   1665
         Width           =   2295
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "中药配方(&1)"
         Height          =   350
         Index           =   1
         Left            =   0
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   1335
         Width           =   2295
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "诊疗项目(&0)"
         Height          =   350
         Index           =   0
         Left            =   0
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   15
         Width           =   2295
      End
      Begin MSComctlLib.TreeView tvwClass 
         Height          =   4005
         Left            =   45
         TabIndex        =   5
         Tag             =   "1000"
         Top             =   2055
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   7064
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "imgList"
         Appearance      =   0
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   150
      Top             =   7170
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   28
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":08CA
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":0E64
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":13FE
            Key             =   "检验U"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":1998
            Key             =   "检验S"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":1F32
            Key             =   "检查U"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":24CC
            Key             =   "检查S"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":2A66
            Key             =   "处置U"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":3000
            Key             =   "处置S"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":359A
            Key             =   "手术U"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":3B34
            Key             =   "手术S"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":40CE
            Key             =   "麻醉U"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":4668
            Key             =   "麻醉S"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":4C02
            Key             =   "护理U"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":519C
            Key             =   "护理S"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":5736
            Key             =   "膳食U"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":5CD0
            Key             =   "膳食S"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":626A
            Key             =   "输血U"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":6804
            Key             =   "输血S"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":6D9E
            Key             =   "输氧U"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":7338
            Key             =   "输氧S"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":78D2
            Key             =   "其他U"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":7E6C
            Key             =   "其他S"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":8406
            Key             =   "成药U"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":89A0
            Key             =   "成药S"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":8F3A
            Key             =   "草药U"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":94D4
            Key             =   "草药S"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":9A6E
            Key             =   "方案U"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":A008
            Key             =   "方案S"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   4125
      Left            =   4440
      TabIndex        =   1
      Top             =   960
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   7276
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
      Top             =   8355
      Width           =   12450
      _ExtentX        =   21960
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmClinicLists.frx":A5A2
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16880
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
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":AE34
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":B04E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":B268
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":B482
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":B69C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":B8B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":BAD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":BCEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":BF04
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":C11E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":C33E
            Key             =   "Quit"
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
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":C55E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":C77E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":C99E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":CBB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":CDD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":CFEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":D206
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":D420
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":D63A
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":D854
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLists.frx":DA74
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin SysInfoLib.SysInfo SysInfo 
      Left            =   210
      Top             =   7245
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin TabDlg.SSTab tabContent 
      Height          =   2820
      HelpContextID   =   1
      Left            =   2760
      TabIndex        =   11
      Top             =   5505
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   4974
      _Version        =   393216
      Style           =   1
      Tabs            =   9
      Tab             =   4
      TabsPerRow      =   9
      TabHeight       =   520
      WordWrap        =   0   'False
      OLEDropMode     =   1
      TabCaption(0)   =   "执行科室(&S)"
      TabPicture(0)   =   "frmClinicLists.frx":DC94
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "hgd定向执行"
      Tab(0).Control(1)=   "fraSubInfo(0)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "收费对照(&C)"
      TabPicture(1)   =   "frmClinicLists.frx":DCB0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraSubInfo(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "检验指标(&L)"
      TabPicture(2)   =   "frmClinicLists.frx":DCCC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraSubInfo(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "可选部位(&P)"
      TabPicture(3)   =   "frmClinicLists.frx":DCE8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraSubInfo(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "用法用量(&U)"
      TabPicture(4)   =   "frmClinicLists.frx":DD04
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "fraSubInfo(4)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "配伍禁忌(T)"
      TabPicture(5)   =   "frmClinicLists.frx":DD20
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fraSubInfo(5)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "配方组成(&M)"
      TabPicture(6)   =   "frmClinicLists.frx":DD3C
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fraSubInfo(6)"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "成套方案(&M)"
      TabPicture(7)   =   "frmClinicLists.frx":DD58
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "fraSubInfo(7)"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "应用参考(&R)"
      TabPicture(8)   =   "frmClinicLists.frx":DD74
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "fraSubInfo(8)"
      Tab(8).ControlCount=   1
      Begin VB.Frame fraSubInfo 
         Height          =   3195
         Index           =   8
         Left            =   -74250
         TabIndex        =   20
         Top             =   495
         Width           =   5115
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdRefer 
            Height          =   2880
            Left            =   150
            TabIndex        =   35
            Top             =   195
            Width           =   5850
            _ExtentX        =   10319
            _ExtentY        =   5080
            _Version        =   393216
            BackColor       =   -2147483628
            Rows            =   5
            Cols            =   4
            FixedRows       =   0
            BackColorBkg    =   -2147483628
            GridColor       =   -2147483628
            GridColorFixed  =   16777215
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            GridLines       =   0
            GridLinesFixed  =   0
            ScrollBars      =   2
            MergeCells      =   1
            AllowUserResizing=   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   4
         End
      End
      Begin VB.Frame fraSubInfo 
         Height          =   3195
         Index           =   7
         Left            =   -73980
         TabIndex        =   18
         Top             =   300
         Width           =   8010
         Begin VSFlex8Ctl.VSFlexGrid vsScheme 
            Height          =   2625
            Left            =   165
            TabIndex        =   19
            Top             =   225
            Width           =   7035
            _cx             =   12409
            _cy             =   4630
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
            BackColorSel    =   12632256
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   22
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   250
            RowHeightMax    =   2000
            ColWidthMin     =   0
            ColWidthMax     =   5000
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmClinicLists.frx":DD90
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
            AutoSizeMode    =   1
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
            OwnerDraw       =   1
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   -1  'True
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
      Begin VB.Frame fraSubInfo 
         Height          =   2430
         Index           =   6
         Left            =   -74835
         TabIndex        =   17
         Top             =   360
         Width           =   6570
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdRecipe 
            Height          =   1710
            Left            =   195
            TabIndex        =   33
            Top             =   315
            Width           =   3960
            _ExtentX        =   6985
            _ExtentY        =   3016
            _Version        =   393216
            FixedCols       =   0
            AllowUserResizing=   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fraSubInfo 
         Height          =   3195
         Index           =   5
         Left            =   -74940
         TabIndex        =   16
         Top             =   405
         Width           =   8010
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdTabu 
            Height          =   2175
            Left            =   120
            TabIndex        =   32
            Top             =   150
            Width           =   7260
            _ExtentX        =   12806
            _ExtentY        =   3836
            _Version        =   393216
            FixedCols       =   0
            AllowUserResizing=   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fraSubInfo 
         Height          =   3195
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   525
         Width           =   8010
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdUsage 
            Height          =   2175
            Left            =   210
            TabIndex        =   34
            Top             =   45
            Width           =   7260
            _ExtentX        =   12806
            _ExtentY        =   3836
            _Version        =   393216
            FixedCols       =   0
            AllowUserResizing=   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fraSubInfo 
         Height          =   2190
         Index           =   3
         Left            =   -74835
         TabIndex        =   14
         Top             =   450
         Width           =   7710
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdPart 
            Height          =   2175
            Left            =   195
            TabIndex        =   31
            Top             =   375
            Width           =   7260
            _ExtentX        =   12806
            _ExtentY        =   3836
            _Version        =   393216
            FixedCols       =   0
            AllowUserResizing=   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fraSubInfo 
         Height          =   2265
         Index           =   2
         Left            =   -74895
         TabIndex        =   13
         Top             =   345
         Width           =   7740
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdLabs 
            Height          =   2175
            Left            =   90
            TabIndex        =   30
            Top             =   180
            Width           =   7260
            _ExtentX        =   12806
            _ExtentY        =   3836
            _Version        =   393216
            FixedCols       =   0
            AllowUserResizing=   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fraSubInfo 
         BorderStyle     =   0  'None
         Height          =   2355
         Index           =   1
         Left            =   -74895
         TabIndex        =   12
         Top             =   465
         Width           =   7845
         Begin VSFlex8Ctl.VSFlexGrid vsfExse 
            Height          =   1425
            Left            =   1095
            TabIndex        =   42
            Top             =   195
            Width           =   5955
            _cx             =   10504
            _cy             =   2514
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
            BackColorSel    =   12632256
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   22
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   250
            RowHeightMax    =   2000
            ColWidthMin     =   0
            ColWidthMax     =   5000
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
            AutoSizeMode    =   1
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
            OwnerDraw       =   1
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   -1  'True
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgd定向执行 
         Height          =   1125
         Left            =   -71325
         TabIndex        =   37
         Top             =   1230
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   1984
         _Version        =   393216
         FixedCols       =   0
         ScrollBars      =   2
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame fraSubInfo 
         Enabled         =   0   'False
         Height          =   2205
         Index           =   0
         Left            =   -74910
         TabIndex        =   21
         Top             =   345
         Width           =   7890
         Begin VB.OptionButton opt执行部门 
            Caption         =   "开单人所在科室(&6)"
            Height          =   195
            Index           =   6
            Left            =   1035
            TabIndex        =   40
            Top             =   1830
            Width           =   1860
         End
         Begin VB.OptionButton opt执行部门 
            Caption         =   "医院外执行(&5)"
            Height          =   195
            Index           =   5
            Left            =   1035
            TabIndex        =   27
            Top             =   1590
            Width           =   2250
         End
         Begin VB.OptionButton opt执行部门 
            Caption         =   "指定科室执行(&4)"
            Height          =   195
            Index           =   4
            Left            =   1035
            TabIndex        =   26
            Top             =   1350
            Width           =   2250
         End
         Begin VB.OptionButton opt执行部门 
            Caption         =   "操作员所在科室(&3)"
            Height          =   195
            Index           =   3
            Left            =   1035
            TabIndex        =   25
            Top             =   1110
            Width           =   2250
         End
         Begin VB.OptionButton opt执行部门 
            Caption         =   "由病人病区执行(&2)"
            Height          =   195
            Index           =   2
            Left            =   1035
            TabIndex        =   24
            Top             =   870
            Width           =   2250
         End
         Begin VB.OptionButton opt执行部门 
            Caption         =   "由病人科室执行(&1)"
            Height          =   195
            Index           =   1
            Left            =   1035
            TabIndex        =   23
            Top             =   630
            Width           =   2250
         End
         Begin VB.OptionButton opt执行部门 
            Caption         =   "不跟踪执行的叮嘱(&0)"
            Height          =   195
            Index           =   0
            Left            =   1035
            TabIndex        =   22
            Top             =   390
            Value           =   -1  'True
            Width           =   2250
         End
         Begin VB.Label lblExcute 
            AutoSize        =   -1  'True
            Caption         =   "执行科室："
            Height          =   180
            Left            =   150
            TabIndex        =   39
            Top             =   390
            Width           =   900
         End
         Begin VB.Label lblUseBill 
            AutoSize        =   -1  'True
            Caption         =   "诊疗单据："
            Height          =   180
            Left            =   150
            TabIndex        =   38
            Top             =   165
            Width           =   900
         End
         Begin VB.Label lbl定向执行 
            AutoSize        =   -1  'True
            Caption         =   "以下科室开单分别一般由指定科室执行(&L)："
            Height          =   180
            Left            =   3570
            TabIndex        =   29
            Top             =   645
            Width           =   3510
         End
         Begin VB.Label lbl常规执行 
            AutoSize        =   -1  'True
            Caption         =   "一般门诊由              执行；"
            Height          =   180
            Left            =   3570
            TabIndex        =   28
            Top             =   375
            Width           =   2700
         End
      End
   End
   Begin VB.Label lblScale 
      AutoSize        =   -1  'True
      Caption         =   "比例尺寸"
      Height          =   180
      Left            =   9330
      TabIndex        =   36
      Top             =   4500
      Visible         =   0   'False
      Width           =   1185
      WordWrap        =   -1  'True
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
      End
      Begin VB.Menu mnuFileSpt2 
         Caption         =   "-"
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
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "项目(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "新增(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuAddcopy 
         Caption         =   "复制新增(&C)"
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditRefer 
         Caption         =   "应用参考(&R)..."
         Shortcut        =   ^R
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditExse 
         Caption         =   "收费对照(&E)..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuEditLabs 
         Caption         =   "检验指标(&L)..."
         Shortcut        =   ^L
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditGather 
         Caption         =   "采集方式(&G)..."
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuEditSample 
         Caption         =   "标本对照(&P)..."
         Shortcut        =   ^S
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStart 
         Caption         =   "启用(&R)"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "停用(&S)"
      End
      Begin VB.Menu mnuEditSpt3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditRepellent 
         Caption         =   "排斥关系(&N)"
      End
      Begin VB.Menu mnuEditBill 
         Caption         =   "对应单据(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuEditSpt4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditImport 
         Caption         =   "项目导入(&I)"
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
         Caption         =   "显示所有下级(&H)"
      End
      Begin VB.Menu mnuViewStoped 
         Caption         =   "显示停用(&C)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "查找(&F)..."
         Shortcut        =   ^F
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
Attribute VB_Name = "frmClinicLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint范围 As Integer '成套方案的可使用场合，1-门诊,2-住院,3-门诊和住院
Private mlngMode As Long
Private mstrPrivs As String       '用户具有本程序的具体权限
Private mbyt中药味数 As Byte

Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim intCount As Integer, intRow As Integer, intCol As Integer
Dim strTemp As String
Private mblnPACSInterface As Boolean        '启用影像信息系统接口

Private Const conTab执行科室 As Integer = 0
Private Const conTab收费对照 As Integer = 1
Private Const conTab检验指标 As Integer = 2
Private Const conTab检查部位 As Integer = 3
Private Const conTab用法用量 As Integer = 4
Private Const conTab配伍禁忌 As Integer = 5
Private Const conTab配方组成 As Integer = 6
Private Const conTab成套方案 As Integer = 7
Private Const conTab应用参考 As Integer = 8

Private Enum SelectKind
    SK_诊疗项目 = "0"
    SK_中药配方 = "1"
    SK_成套方案 = "2"
End Enum

Private Enum COL成套方案
    col期效 = 0
    col内容 = 1
    col总量 = 2
    col总量单位 = 3
    col单量 = 4
    col单位 = 5
    col天数 = 6
    col频次 = 7
    col用法 = 8
    col嘱托 = 9
    col执行时间 = 10
    col执行科室 = 11
    col执行性质 = 12
    col序号 = 13
    col相关 = 14
    col项目ID = 15
    col类别 = 16
    col收费细目ID = 17
    col标本部位 = 18
    col检查方法 = 19
    col执行标记 = 20 '药品医嘱用于区分 自取药和不取药
    col停用 = 21
End Enum
 
Private mblnStartPriceGrade As Boolean '启用了价格等级
Private mstrPriceGrade As String
Private mstrPriceGradeFields As String

Private Sub InitPriceGrade()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化价格等级
    '编制:刘兴洪
    '日期:2017-07-01 21:37:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String, strTempFileds As String
    Dim i As Long
    mblnStartPriceGrade = zlGetrsPriceGrade(rsTemp)
    mstrPriceGrade = "": mstrPriceGradeFields = ""
    If mblnStartPriceGrade = False Then Exit Sub
    If rsTemp.RecordCount = 0 Then mblnStartPriceGrade = False: Exit Sub
    With rsTemp
        i = 1
        .MoveFirst
        Do While Not .EOF
            mstrPriceGrade = mstrPriceGrade & "," & !名称
            strTempFileds = strTempFileds & ",sum(decode(P.价格等级,'" & !名称 & "',P.现价, -1*NULL))  as   A" & i
            i = i + 1
            .MoveNext
        Loop
        .MoveFirst
    End With
    If mstrPriceGrade <> "" Then mstrPriceGrade = Mid(mstrPriceGrade, 2)
    mstrPriceGradeFields = strTempFileds
End Sub



Public Sub ShowMeWithScheme(frmMain As Object, ByVal int范围 As Integer)
    mint范围 = int范围
    
    On Error Resume Next
    Me.Show , frmMain
    Me.Caption = "成套方案管理"
End Sub

Private Sub cmdKind_Click(Index As Integer)
    Dim intCount As Integer
    
    Call SaveListViewState(Me.lvwItems, Me.Name & Val(Me.tvwClass.Tag), Me.lvwItems.View)
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        If intCount <= Index Then
            Me.cmdKind(intCount).Tag = SK_诊疗项目
        Else
            Me.cmdKind(intCount).Tag = SK_中药配方
        End If
    Next
    
    '设置导入的菜单属性
    Me.mnuEditImport.Enabled = (Index = SK_诊疗项目)
    '装数据并调整界面
    If Me.lvwItems.Visible Then
        Call picClass_Resize
        Me.tvwClass.SetFocus
    End If
    If Val(Me.tvwClass.Tag) <> Index Then
        Me.tvwClass.Tag = Index
        Call zlRefClasses
    End If
    Me.mnuClass.Enabled = (tvwClass.Tag = SK_诊疗项目 And InStr(1, mstrPrivs, "诊疗项目编辑") > 0) Or _
                            (tvwClass.Tag = SK_中药配方 And InStr(1, mstrPrivs, "中药配方编辑") > 0) Or _
                            (tvwClass.Tag = SK_成套方案 And InStr(1, mstrPrivs, "成套方案编辑") > 0)
    
    Me.mnuEditAdd.Enabled = Me.mnuClass.Enabled
    Me.mnuAddcopy.Enabled = Me.mnuClass.Enabled
    Me.mnuEditModify.Enabled = Me.mnuClass.Enabled
    Me.mnuEditDelete.Enabled = Me.mnuClass.Enabled
    Me.mnuEditLabs.Enabled = Me.mnuClass.Enabled
    Me.mnuEditGather.Enabled = Me.mnuClass.Enabled
    Me.mnuEditStart.Enabled = Me.mnuClass.Enabled
    Me.mnuEditStop.Enabled = Me.mnuClass.Enabled
    Me.mnuEditRepellent.Tag = ""
    Me.mnuEditBill.Enabled = Me.mnuClass.Enabled
    Me.tlbThis.Buttons("Class").Enabled = Me.mnuClass.Enabled
    Me.tlbThis.Buttons("Add").Enabled = Me.mnuClass.Enabled
    Me.tlbThis.Buttons("Modify").Enabled = Me.mnuClass.Enabled
    Me.tlbThis.Buttons("Delete").Enabled = Me.mnuClass.Enabled
    Me.tlbThis.Buttons("Start").Enabled = Me.mnuClass.Enabled
    Me.tlbThis.Buttons("Stop").Enabled = Me.mnuClass.Enabled
End Sub

Private Sub clbThis_Resize()
    Me.clbThis.Bands(1).MinHeight = Me.tlbThis.Height
    Me.clbThis.Refresh
    Call Form_Resize
End Sub

Private Sub Form_Activate()
    Me.lvwItems.Visible = True
End Sub

Private Sub Form_Load()
    '界面恢复
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    
    mblnPACSInterface = (Val(zlDatabase.GetPara(255, glngSys, , "0")) = 1)
    Call InitPriceGrade
    
    
    If mint范围 = 0 Then mint范围 = 3 '直接通过诊疗项目管理进入时，缺省为3
    
    Call RestoreWinState(Me, App.ProductName)
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    If Val(zlDatabase.GetPara("使用个性化风格", , , True)) = 1 Then
        strTemp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\分割", "横向", "0")
        If strTemp <> "0" Then
            Me.picVBar.Left = CLng(strTemp)
        End If
        strTemp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\分割", "纵向", "0")
        If strTemp <> "0" Then
            Me.picHBar.Top = CLng(strTemp)
        End If
    End If
    
    Me.mnuViewStoped.Checked = (Val(zlDatabase.GetPara("显示停用项目", glngSys, 1054, 0)) = 1)
    With Me.hgdRefer
        .ColWidth(0) = 0
        .ColWidth(1) = Me.TextWidth("空格")
        .ColWidth(2) = .Width - .ColWidth(1) - Me.SysInfo.ScrollBarSize - 15
        .ColWidth(3) = 600
    End With
    
    '可直接通过菜单进行的权限控制
    If InStr(1, mstrPrivs, "诊疗项目编辑") = 0 And _
                            InStr(1, mstrPrivs, "中药配方编辑") = 0 And _
                            InStr(1, mstrPrivs, "成套方案编辑") = 0 Then
        Me.mnuClass.Enabled = False
        Me.mnuEditAdd.Enabled = False
        Me.mnuAddcopy.Enabled = False
        Me.mnuEditModify.Enabled = False
        Me.mnuEditDelete.Enabled = False
        Me.mnuEditLabs.Enabled = False
        Me.mnuEditGather.Enabled = False
        'Me.mnuEditSample.Enabled = False
        'Me.mnuEditExams.Enabled = False
        Me.mnuEditStart.Enabled = False
        Me.mnuEditStop.Enabled = False
        Me.mnuEditRepellent.Tag = ""
        Me.mnuEditBill.Enabled = False
        Me.tlbThis.Buttons("Class").Enabled = False
        Me.tlbThis.Buttons("Add").Enabled = False
        Me.tlbThis.Buttons("Modify").Enabled = False
        Me.tlbThis.Buttons("Delete").Enabled = False
        Me.tlbThis.Buttons("Start").Enabled = False
        Me.tlbThis.Buttons("Stop").Enabled = False
    Else
        Me.mnuEditRepellent.Tag = 1
    End If
    If InStr(1, mstrPrivs, "收费设置") = 0 Then
        Me.mnuEditExse.Enabled = False
    End If
    If InStr(1, mstrPrivs, "参考编辑") = 0 Then
'        Me.mnuEditRefer.Enabled = False
    End If
    If InStr(1, mstrPrivs, "项目导入") = 0 Then
        Me.mnuEditSpt4.Visible = False
        Me.mnuEditImport.Visible = False
    End If
    
    If InStr(mstrPrivs, "管理诊疗项目") = 0 And InStr(mstrPrivs, "管理中药配方") = 0 And InStr(mstrPrivs, "管理成套方案") = 0 Then
        MsgBox "你没有管理任何项目内容的权限，请与系统管理员联系。", vbInformation, gstrSysName
        Unload Me: Exit Sub
    Else
        If InStr(mstrPrivs, "管理诊疗项目") = 0 Then
            cmdKind(0).Visible = False
        End If
        If InStr(mstrPrivs, "管理中药配方") = 0 Then
            cmdKind(1).Visible = False
        End If
        If InStr(mstrPrivs, "管理成套方案") = 0 Then
            cmdKind(2).Visible = False
        End If
        
        If InStr(mstrPrivs, "管理诊疗项目") > 0 Then
            Call cmdKind_Click(0)
        ElseIf InStr(mstrPrivs, "管理中药配方") > 0 Then
            Call cmdKind_Click(1)
        ElseIf InStr(mstrPrivs, "管理成套方案") > 0 Then
            Call cmdKind_Click(2)
        End If
    End If
    
    '初始化新网RIS接口
    If mblnPACSInterface Then
        Call IniRIS
    End If
End Sub

Private Sub Form_Resize()
    Dim lngTools As Single, lngStatus As Single
    Dim i As Integer
    
    If WindowState = 1 Then Exit Sub
    lngTools = IIf(Me.clbThis.Visible, Me.clbThis.Height, 0)
    lngStatus = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    err = 0: On Error Resume Next
    
    With Me.picVBar
        .Top = lngTools
        .Height = Me.ScaleHeight - picClass.Top - lngStatus
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
    
    For intCount = 0 To Me.tabContent.Tabs - 1
        With Me.fraSubInfo(intCount)
            .Left = 90
            .Top = 325
            .Width = Me.tabContent.Width - .Left * 2
            .Height = Me.tabContent.Height - .Top - 90
        End With
    Next
    With Me.hgd定向执行
'        .Visible = Me.fraSubInfo(0).Visible
        .Left = Me.fraSubInfo(0).Left + Me.lbl定向执行.Left
        .Width = Me.fraSubInfo(0).Left + Me.fraSubInfo(0).Width - .Left - 100
        .Top = Me.fraSubInfo(0).Top + Me.lbl定向执行.Top + Me.lbl定向执行.Height + 45
        .Height = Me.fraSubInfo(0).Top + Me.fraSubInfo(0).Height - .Top - 100
    End With
    With vsfExse
        .Left = 30: .Top = 30
        .Width = fraSubInfo(1).Width - 60
        .Height = fraSubInfo(1).Height - 60
    End With
    
    With Me.hgdLabs
        .Left = 0: .Top = 90: .Width = Me.fraSubInfo(2).Width: .Height = Me.fraSubInfo(2).Height - .Top
    End With
    With Me.hgdPart
        .Left = 0: .Top = 90: .Width = Me.fraSubInfo(3).Width: .Height = Me.fraSubInfo(3).Height - .Top
    End With
    With Me.hgdUsage
        .Left = 0: .Top = 90: .Width = Me.fraSubInfo(4).Width: .Height = Me.fraSubInfo(4).Height - .Top
    End With
    With Me.hgdTabu
        .Left = 0: .Top = 90: .Width = Me.fraSubInfo(5).Width: .Height = Me.fraSubInfo(5).Height - .Top
    End With
    With Me.hgdRecipe
        .Left = 0: .Top = 90: .Width = Me.fraSubInfo(6).Width: .Height = Me.fraSubInfo(6).Height - .Top
    End With
    With Me.vsScheme '成套方案(ByZT)
        .Left = 0: .Top = 90: .Width = Me.fraSubInfo(7).Width: .Height = Me.fraSubInfo(7).Height - .Top
    End With

    With Me.hgdRefer
        .Left = 0: .Top = 90: .Width = Me.fraSubInfo(6).Width: .Height = Me.fraSubInfo(6).Height - .Top
        .Redraw = False
        .ColWidth(0) = 0
        .ColWidth(1) = Me.TextWidth("空格")
        .ColWidth(2) = .Width - .ColWidth(1) - Me.SysInfo.ScrollBarSize - 15
        .ColWidth(3) = 600
        Call zlGrdRowHeight
        .Redraw = True
    End With
    clbThis.Bands(1).Width = Me.Width - 2000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    Call SaveListViewState(Me.lvwItems, Me.Name & Val(Me.tvwClass.Tag), Me.lvwItems.View)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\分割", "横向", Me.picVBar.Left)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\分割", "纵向", Me.picHBar.Top)
    
    Call zlDatabase.SetPara("显示停用项目", IIf(Me.mnuViewStoped.Checked, 1, 0), glngSys, 1054)
    
    mint范围 = 3 '直接通过诊疗项目管理进入时，缺省为3
    
    If Not gobjRIS Is Nothing Then
        Set gobjRIS = Nothing
    End If
End Sub

Private Sub mnuAddcopy_Click()
    Dim blnOk As Boolean
    
    If Me.tvwClass.SelectedItem Is Nothing Then MsgBox "尚未设置分类,不能增删项目！", vbExclamation, gstrSysName: Exit Sub
    If Val(Me.tvwClass.Tag) = 0 Then
        If Me.lvwItems.SelectedItem Is Nothing Then
            MsgBox "请选择需要复制的项目！", vbInformation, gstrSysName
        Else
            blnOk = frmClinicItem.ShowMe(Me, 3, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2))
        End If
    End If
    If blnOk Then Call zlRefRecords
End Sub

Private Sub tlbThis_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    If tvwClass.Tag <> SK_诊疗项目 Then
        Button.ButtonMenus(2).Visible = False
    Else
        Button.ButtonMenus(2).Visible = True
    End If
End Sub

Private Sub vsfExse_DblClick()
    Dim i As Integer
    Dim strIDS As String
    
    If Val(Me.tvwClass.Tag) <> 0 Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    For i = Me.lvwItems.SelectedItem.Index + 1 To lvwItems.ListItems.Count
        strIDS = strIDS & Mid(Me.lvwItems.ListItems(i).Key, 2) & ","
    Next
    
    Call frmClinicExse.ShowMe(Me, True, Mid(Me.lvwItems.SelectedItem.Key, 2), strIDS)
End Sub

Private Sub hgdLabs_DblClick()
'    If Val(Me.tvwClass.Tag) <> 0 Then Exit Sub
'    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
'    Call frmClinicLabs.ShowME(Me, False, Mid(Me.lvwItems.SelectedItem.Key, 2))
End Sub

Private Sub hgdPart_DblClick()
    If Val(Me.tvwClass.Tag) <> 0 Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call frmClinicPart.ShowMe(Me, False, Mid(Me.lvwItems.SelectedItem.Key, 2))
End Sub

Private Sub hgdRecipe_DblClick()
    If Val(Me.tvwClass.Tag) <> 0 Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call frmMediRecipe.ShowMe(Me, 2, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2))
End Sub

Private Sub hgdRefer_DblClick()
    If Me.mnuEditRefer.Enabled = False Then Exit Sub
End Sub

Private Sub mnuEditImport_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then MsgBox "尚未设置分类,不能导入项目！", vbExclamation, gstrSysName: Exit Sub
    With frmClinicLoad
        .Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
        .Show 1, Me
    End With
    Call zlRefRecords
End Sub

'Private Sub mnuEditSample_Click()
'    '2007-04-17 去掉标本对照功能
''    If Me.lvwItems.ListItems.Count > 0 Then
''        Call frmClinicVerifySample.ShowMe(Me, Mid(Me.lvwItems.SelectedItem.Key, 2))
''    End If
'End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    '默认参数：分类=分类id，项目=项目id，类别=诊疗类别名称
    Dim lng分类id As Long
    Dim lng项目id As Long
    Dim str类别 As String
    
    If Not Me.tvwClass.SelectedItem Is Nothing Then
        lng分类id = Mid(Me.tvwClass.SelectedItem.Key, 2)
    End If
    
    If Not Me.lvwItems.SelectedItem Is Nothing Then
        lng项目id = Mid(Me.lvwItems.SelectedItem.Key, 2)
        str类别 = Me.lvwItems.SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("_类别").Index - 1)
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "分类=" & IIf(lng分类id = 0, "", lng分类id), _
        "项目=" & IIf(lng项目id = 0, "", lng项目id), _
        "类别=" & str类别)
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

Private Sub tabCharge_Click(PreviousTab As Integer)
    vsfExse.ZOrder 0
End Sub

Private Sub txtFind_GotFocus()
    Call zlControl.TxtSelAll(txtFind)
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql As String
    Dim rsTmp As Recordset
    Dim strSQLTmp As String
    Dim vRect As RECT, blnCancel As Boolean
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If txtFind.Text = "" Then Exit Sub
    If zlCommFun.IsCharChinese(txtFind.Text) Then
        strSQLTmp = " And Upper(Nvl(b.名称, a.名称)) Like [1]"
    ElseIf IsNumeric(txtFind.Text) Then
        strSQLTmp = " And a.编码 Like [2]"
    Else
        strSQLTmp = " And (Upper(Nvl(b.名称, a.名称)) Like [1] Or b.简码 Like [3])"
    End If
    
    On Error GoTo ErrHandle
    strSql = "Select Distinct a.Id, a.类别, Nvl(b.名称, a.名称) As 名称, a.编码, b.简码, c.名称 As 分类, a.分类id, a.撤档时间" & vbNewLine & _
            "From (Select ID, 类别, 分类id, 名称, 编码, 撤档时间" & vbNewLine & _
            "       From 诊疗项目目录" & vbNewLine & _
            "       Where 类别 Not In ('4', '5', '6', '7') And 类别 >= 'A' And" & vbNewLine & _
            "             (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)) A," & vbNewLine & _
            "     (Select Distinct a.诊疗项目id, a.名称, a.简码 As 拼音码, b.简码 As 五笔码, a.简码 || '/' || b.简码 As 简码" & vbNewLine & _
            "       From 诊疗项目别名 A, 诊疗项目别名 B" & vbNewLine & _
            "       Where a.诊疗项目id = b.诊疗项目id And a.码类 = 1 And b.码类 = 2" & _
            IIf(zlCommFun.IsCharChinese(txtFind.Text), " And Upper(a.名称) Like [1] ", "") & _
            " And a.性质 = 1 And b.性质 = 1) B," & vbNewLine & _
            "     诊疗分类目录 C" & vbNewLine & _
            "Where a.分类id = c.Id(+) And a.Id = b.诊疗项目id(+) And c.名称 Is Not Null And C.类型 In (4,5,6) " & _
            strSQLTmp
    
    vRect = zlControl.GetControlRect(txtFind.hwnd)
    If vRect.Left + 7000 > Screen.Width Then vRect.Left = Screen.Width - 7000
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "收费细目选择", False, "", "", False, False, True, _
                        vRect.Left, vRect.Top, txtFind.Height, blnCancel, False, True, IIf(gstrMatch = "", "", "%") & txtFind.Text & "%", txtFind.Text & "%", IIf(gstrMatch = "", "", "%") & UCase(txtFind.Text) & "%")
    If blnCancel = True Then Exit Sub
    If Not rsTmp Is Nothing Then
        Call FindLocate(rsTmp)
    Else
        MsgBox "没有找到您所查找的收费项目。", vbInformation, Me.Caption
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FindLocate(ByVal rsTmp As Recordset)
    Dim strkey As String
    Dim strItemKey As String
    
    '81291--查找是基于所有分类进行查找，因此不需要判断当前分类下是否是子项
'    If lvwItems.SelectedItem Is Nothing Then Exit Sub

    On Error Resume Next
    With lvwItems.SelectedItem
        strkey = "_" & IIf(IsNull(rsTmp("分类ID")), "", rsTmp("分类ID"))
        strItemKey = "_" & rsTmp("id")
        If .SubItems(3) <> "未分类" Then
            Me.tvwClass.Nodes(strkey).Selected = True
            Me.tvwClass.Nodes(strkey).EnsureVisible
            Me.tvwClass_NodeClick Me.tvwClass.SelectedItem
            err.Clear
            Me.lvwItems.ListItems(strItemKey).Selected = True
            Me.lvwItems.ListItems(strItemKey).EnsureVisible
            If err.Number = 35601 Then
                MsgBox "你找到的这条记录可能已被删除或停用，请刷新列表。", vbInformation, gstrSysName
                err.Clear
                Exit Sub
            End If
            Me.lvwItems_ItemClick Me.lvwItems.SelectedItem
        Else
            Me.tvwClass.Nodes("Root").Selected = True
            Me.tvwClass.Nodes(strkey).EnsureVisible
            Me.tvwClass_NodeClick Me.tvwClass.SelectedItem
            err.Clear
            Me.lvwItems.ListItems(strItemKey).Selected = True
            Me.lvwItems.ListItems(strItemKey).EnsureVisible
            If err.Number = 35601 Then
                MsgBox "你找到的这条记录可能已被删除或停用，请刷新列表。", vbInformation, gstrSysName
                err.Clear
                Exit Sub
            End If
            Me.lvwItems_ItemClick Me.lvwItems.SelectedItem
        End If
    End With
    err.Clear
End Sub

 

Private Sub vsScheme_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= vsScheme.FixedRows And NewCol >= vsScheme.FixedCols Then
        If NewRow <> OldRow Then
            vsScheme.ForeColorSel = vsScheme.CellForeColor
        End If
    End If
End Sub

Private Sub mnuEditGather_Click()
    If Val(Me.tvwClass.Tag) <> 0 Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    If Me.lvwItems.SelectedItem.Tag <> "C" Then Exit Sub
    Call frmLabsUsage.ShowMe(Me, Not Me.lvwItems.SelectedItem.Icon = "诊疗S", Mid(Me.lvwItems.SelectedItem.Key, 2))
    Call zlRefRecords(Mid(Me.lvwItems.SelectedItem.Key, 2))
    
    If Not Me.lvwItems.SelectedItem Is Nothing Then
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    End If
End Sub

Private Sub vsScheme_DblClick()
    '查阅成套方案(ByZT)
    If Val(Me.tvwClass.Tag) <> 2 Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call frmClinicScheme.ShowMe(Me, mstrPrivs, 1, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2), mint范围)
End Sub

Private Sub hgdUsage_DblClick()
    If Val(Me.tvwClass.Tag) <> 0 Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call frmMediUsage.ShowMe(Me, False, Mid(Me.lvwItems.SelectedItem.Key, 2))
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
    Select Case Val(Me.tvwClass.Tag)
    Case 0
        Call frmClinicItem.ShowMe(Me, 2, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2))
    Case 1
        Call frmMediRecipe.ShowMe(Me, 2, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2))
    Case 2
        '查阅成套方案(ByZT)
        Call frmClinicScheme.ShowMe(Me, mstrPrivs, 2, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2), mint范围)
    End Select
End Sub

Public Sub lvwItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim i As Long, j As Long
    Dim iRow As Integer, iCol As Integer
    Dim lngForeColor As Long
    Dim strTmp As String
    Dim intIndex As Integer
    Dim lngCol As Long
    Dim lngRow As Long
    '------------------------------------------------
    '清理详细信息显示区
    Call zlClearDetail
    
    err = 0: On Error GoTo ErrHand
    '------------------------------------------------
    '执行科室显示
    If Val(Me.tvwClass.Tag) = 0 Then
        Me.lblUseBill.Caption = "诊疗单据："
        gstrSql = "Select A.应用场合,B.名称" & _
                " From 病历文件列表 B,病历单据应用 A" & _
                " Where B.ID=A.病历文件id And A.诊疗项目id=[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
        
        With rsTemp
            Do While Not .EOF
                Select Case !应用场合
                Case 1
                    Me.lblUseBill.Caption = Me.lblUseBill.Caption & "门诊采用" & !名称 & "；"
                Case 2
                    Me.lblUseBill.Caption = Me.lblUseBill.Caption & "住院采用" & !名称 & "；"
                Case 4
                    Me.lblUseBill.Caption = Me.lblUseBill.Caption & "体检采用" & !名称 & "；"
                End Select
                
                .MoveNext
            Loop
        End With
        
        gstrSql = "select 执行科室 from 诊疗项目目录 where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
        

        Me.opt执行部门(0).Value = False
        If rsTemp.RecordCount > 0 Then Me.opt执行部门(IIf(IsNull(rsTemp!执行科室), 0, rsTemp!执行科室)).Value = True
        For i = 0 To opt执行部门.Count - 1
            opt执行部门(i).Enabled = opt执行部门(i).Value
        Next
        
        gstrSql = "select R.病人来源,E.ID,E.名称" & _
                " from 诊疗执行科室 R,部门表 E" & _
                " where R.执行科室ID=E.ID and R.病人来源 in (1,2) and R.开单科室id is null and R.诊疗项目ID=[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
            
        With rsTemp
            strTemp = ""
            Do While Not .EOF
                If !病人来源 = 1 Then strTemp = strTemp & "门诊由" & !名称 & "执行；"
                If !病人来源 = 2 Then strTemp = strTemp & "住院由" & !名称 & "执行；"
                .MoveNext
            Loop
        End With
        
        Me.lbl常规执行.Caption = ""
        If strTemp <> "" Then Me.lbl常规执行.Caption = "一般" & strTemp
        
        gstrSql = "select K.名称 as 开单部门名称,E.名称 as 执行部门名称" & _
                " from 诊疗执行科室 R,部门表 K,部门表 E" & _
                " where R.开单科室ID=K.ID(+) and R.执行科室ID=E.ID and nvl(R.病人来源,0)=0 and R.诊疗项目ID=[1] " & _
                " order by e.名称"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
        
        Me.hgd定向执行.Redraw = False
        i = 0
         With rsTemp
            Do While Not .EOF
'                If Me.hgd定向执行.Rows - 1 < .AbsolutePosition Then Me.hgd定向执行.Rows = Me.hgd定向执行.Rows + 1
'                Me.hgd定向执行.TextMatrix(.AbsolutePosition, 0) = !开单部门名称
'                Me.hgd定向执行.TextMatrix(.AbsolutePosition, 1) = !执行部门名称
                
                If strTmp <> !执行部门名称 Then
                    i = i + 1
                    Me.hgd定向执行.Rows = i + 1
                    Me.hgd定向执行.TextMatrix(i, 1) = IIf(IsNull(!开单部门名称), "（所有部门）", !开单部门名称)
                    Me.hgd定向执行.TextMatrix(i, 0) = !执行部门名称
                Else
                    Me.hgd定向执行.TextMatrix(i, 1) = Me.hgd定向执行.TextMatrix(i, 1) & "," & !开单部门名称
                End If
                
                strTmp = !执行部门名称
                
                .MoveNext
            Loop
            Me.hgd定向执行.Redraw = True
        End With
    End If
    
    '------------------------------------------------
    '收费对照显示
    If Val(Me.tvwClass.Tag) = 0 Then
         
        
        
        gstrSql = "" & _
            " Select i.Id, r.检查部位, r.检查方法, r.费用性质, '[' || i.编码 || ']' || i.名称 As 名称, i.规格, i.计算单位," & vbNewLine & _
            "       I.是否变价,sum(decode(P.价格等级,NULL,p.现价,0)) as 缺省价格" & mstrPriceGradeFields & _
            "       ,Nvl(r.收费数量, 0) As 数量, Nvl(r.固有对照, 0) As 固定," & vbNewLine & _
            "       Nvl(r.从属项目, 0) As 从项, Nvl(i.撤档时间, To_Date('3000-01-01', 'YYYY-MM-DD')) As 撤档时间, Nvl(r.收费方式, 0) As 收费方式, r.病人来源," & vbNewLine & _
            "       b.编码 As 适用科室编码, b.名称 As 适用科室名称" & vbNewLine & _
            " From 诊疗收费关系 R, 收费项目目录 I, 收费价目 P, 部门表 B" & vbNewLine & _
            " Where r.收费项目id = i.Id And i.Id = p.收费细目id(+) And r.适用科室id = b.Id(+)  " & vbNewLine & _
            "       And p.执行日期 <= Sysdate And (p.终止日期 Is Null Or p.终止日期 >= Sysdate)  " & vbNewLine & _
            "      And r.诊疗项目id = [1]" & vbNewLine & _
            "Group By i.Id, r.检查部位, r.检查方法, r.费用性质, i.编码, i.名称, i.规格, i.计算单位, i.是否变价, r.收费数量, r.固有对照, r.从属项目, i.撤档时间, r.收费方式," & vbNewLine & _
            "         r.病人来源, b.编码, b.名称" & vbNewLine & _
            "order by r.病人来源,B.名称,nvl(R.从属项目,0)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)), gstrPriceClass)
        With vsfExse
            .Redraw = flexRDNone
            Do While Not rsTemp.EOF
                intIndex = Val(NVL(rsTemp!病人来源))
                i = .Rows - 1: .Rows = .Rows + 1
               .TextMatrix(i, .ColIndex("选择")) = i
               .TextMatrix(i, .ColIndex("部位")) = NVL(rsTemp!检查部位)
               .TextMatrix(i, .ColIndex("方法")) = NVL(rsTemp!检查方法)
               .TextMatrix(i, .ColIndex("项目名")) = NVL(rsTemp!名称)
               .TextMatrix(i, .ColIndex("规格")) = NVL(rsTemp!规格)
               .TextMatrix(i, .ColIndex("单位")) = NVL(rsTemp!计算单位)
               .TextMatrix(i, .ColIndex("价格")) = IIf(Val(NVL(rsTemp!是否变价)) = 1, "变价", Val(NVL(rsTemp!缺省价格)))
               .TextMatrix(i, .ColIndex("数量")) = FormatEx(Format(rsTemp!数量, "0.00000"), 5)
               .TextMatrix(i, .ColIndex("固定")) = IIf(rsTemp!固定 = 0, "", "√")
               .TextMatrix(i, .ColIndex("从项")) = IIf(rsTemp!从项 = 0, "", "√")
               .TextMatrix(i, .ColIndex("加收")) = IIf(0 + rsTemp!费用性质 = 1, "√", "")
               .TextMatrix(i, .ColIndex("状态")) = IIf(Format(rsTemp!撤档时间, "YYYY-MM-DD") <> "3000-01-01", "停用", "")
                
                Select Case rsTemp!收费方式
                Case 0
                   .TextMatrix(i, .ColIndex("收费方式")) = "0-正常收取"
                Case 1
                   .TextMatrix(i, .ColIndex("收费方式")) = "1-检验试管费用"
                Case 2
                   .TextMatrix(i, .ColIndex("收费方式")) = "2-一次发送只收取一次"
                Case 3
                   .TextMatrix(i, .ColIndex("收费方式")) = "3-当天只收取一次"
                Case 4
                   .TextMatrix(i, .ColIndex("收费方式")) = "4-当天未执行收取一次"
                Case 5
                   .TextMatrix(i, .ColIndex("收费方式")) = "5-当天只收取一次，排斥其他项目"
                Case 6
                   .TextMatrix(i, .ColIndex("收费方式")) = "6-当天未执行收取一次，排斥其他项目"
                Case 7
                   .TextMatrix(i, .ColIndex("收费方式")) = "7-每天首次不收取"
                Case 9
                   .TextMatrix(i, .ColIndex("收费方式")) = "9-自定义"
                Case Else
                   .TextMatrix(i, .ColIndex("收费方式")) = "0-正常收取"
                End Select
                
                Select Case Val(NVL(rsTemp!病人来源))
                Case 0: .TextMatrix(i, .ColIndex("适用场合")) = "所有科室"
                Case 1: .TextMatrix(i, .ColIndex("适用场合")) = "门诊科室"
                Case 2: .TextMatrix(i, .ColIndex("适用场合")) = "住院科室"
                Case 3: .TextMatrix(i, .ColIndex("适用场合")) = "体检科室"
                End Select
                
                If Trim(NVL(rsTemp!适用科室名称)) <> "" Then
                   .TextMatrix(i, .ColIndex("适用科室")) = "" & rsTemp!适用科室名称 & "(" & rsTemp!适用科室编码 & ")"
                End If
                '加载价格等级所对应的价格
                For intCol = 0 To .Cols - 1
                    If Left(CStr(.colData(intCol)), 1) = "A" Then
                        If Val(NVL(rsTemp!是否变价)) = 1 Then
                            .TextMatrix(i, intCol) = "变价"
                        Else
                            If Val(NVL(rsTemp.Fields(CStr(.colData(intCol))))) = 0 Then
                                .TextMatrix(i, intCol) = .TextMatrix(i, .ColIndex("价格"))
                            Else
                                .TextMatrix(i, intCol) = Val(NVL(rsTemp.Fields(CStr(.colData(intCol)))))
                            End If
                        End If
                    End If
                Next
                If Format(rsTemp!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                    lngForeColor = &HFF&
                Else
                    lngForeColor = &H0&
                End If
                iRow = .Row: iCol = .Col
               .Row = i
               .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = lngForeColor
               .Row = iRow: .Col = iCol
                rsTemp.MoveNext
            Loop
            .Redraw = flexRDBuffered
            If .Rows > 2 Then .Rows = .Rows - 1
            .MergeCells = flexMergeFree
            .MergeCol(.ColIndex("适用场合")) = True
        End With
        
    End If
     
    '------------------------------------------------
    '检查部位显示
    If Val(Me.tvwClass.Tag) = 0 And Item.Tag = "D" Then
        gstrSql = "select ID from 诊疗项目目录 I where I.ID=[1] and I.组合项目=1 "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
        
        With rsTemp
            If .EOF Then
                Me.tabContent.TabVisible(conTab检查部位) = False
            Else
                Me.tabContent.TabVisible(conTab检查部位) = True
            End If
        End With
    Else
        Me.tabContent.TabVisible(conTab检查部位) = False
    End If
    If Me.tabContent.TabVisible(conTab检查部位) = True Then
        Me.hgdPart.Redraw = False
        gstrSql = "select I.ID,I.名称 as 名称,I.标本部位" & _
                " from 诊疗项目组合 R,诊疗项目目录 I" & _
                " where R.诊疗项目ID=I.ID and R.诊疗组合ID=[1] " & _
                " order by R.序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
        
        With rsTemp
            Do While Not .EOF
                If Me.hgdPart.Rows - 1 < .AbsolutePosition Then Me.hgdPart.Rows = Me.hgdPart.Rows + 1
                Me.hgdPart.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
                Me.hgdPart.TextMatrix(.AbsolutePosition, 1) = !名称
                Me.hgdPart.TextMatrix(.AbsolutePosition, 2) = !标本部位
                .MoveNext
            Loop
        End With
        Me.hgdPart.Redraw = True
    End If
    '------------------------------------------------
    
    '------------------------------------------------
    '配方组成显示
    If Val(Me.tvwClass.Tag) = 1 Then
        Me.hgdRecipe.Redraw = False
        gstrSql = "Select b.序号, b.诊疗项目Id As 药名id, b.收费细目Id As 规格id, a.名称, c.规格, a.计算单位, b.单次用量, b.医生嘱托 " & vbNewLine & _
            "From 诊疗项目目录 A, 诊疗项目组合 B, 收费项目目录 C " & vbNewLine & _
            "Where a.Id = b.诊疗项目id And b.收费细目id = c.Id(+) And b.诊疗组合id = [1] " & vbNewLine & _
            "Order By b.序号 "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
        
        With rsTemp
            Do While Not .EOF
                If Me.hgdRecipe.Rows - 1 < ((.AbsolutePosition - 1) \ mbyt中药味数) + 1 Then Me.hgdRecipe.Rows = Me.hgdRecipe.Rows + 1
                intCount = (.AbsolutePosition - 1) Mod mbyt中药味数
                Me.hgdRecipe.TextMatrix((.AbsolutePosition - 1) \ mbyt中药味数 + 1, intCount * 6 + 2) = !名称 & IIf(IsNull(!规格), "", "(" & !规格 & ")")
                Me.hgdRecipe.TextMatrix((.AbsolutePosition - 1) \ mbyt中药味数 + 1, intCount * 6 + 3) = IIf(IsNull(!单次用量), 0, !单次用量)
                Me.hgdRecipe.TextMatrix((.AbsolutePosition - 1) \ mbyt中药味数 + 1, intCount * 6 + 4) = IIf(IsNull(!计算单位), "", !计算单位)
                Me.hgdRecipe.TextMatrix((.AbsolutePosition - 1) \ mbyt中药味数 + 1, intCount * 6 + 5) = IIf(IsNull(!医生嘱托), "", !医生嘱托)
                .MoveNext
            Loop
        End With
        
        For lngRow = 1 To hgdRecipe.Rows - 1
            hgdRecipe.Row = lngRow
            For lngCol = 0 To hgdRecipe.Cols - 1
                hgdRecipe.Col = lngCol
                
                If lngCol < 6 Or (lngCol > 12 And lngCol < 20) Then
                    hgdRecipe.CellBackColor = &H8000000F
                End If
            Next
        Next
        
        gstrSql = "select I.名称 ,R.性质,P.名称 as 频率,R.疗程" & _
                " from 诊疗用法用量 R,诊疗项目目录 I,诊疗频率项目 P" & _
                " where R.用法ID=I.ID and R.频次=P.编码(+) and R.项目ID=[1] " & _
                " order by R.性质 desc"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Item.Key, 2)))
            
        With rsTemp
            strTemp = ""
            Do While Not .EOF
                If .AbsolutePosition = 1 Then strTemp = strTemp & Space(3) & IIf(IsNull(!频率), "", !频率)
                strTemp = strTemp & Space(3) & !名称
                .MoveNext
            Loop
            With Me.hgdRecipe
                .Rows = .Rows + 2: .MergeRow(.Rows - 1) = True
                For intCount = 0 To .Cols - 1
                    .TextMatrix(.Rows - 1, intCount) = Trim(strTemp)
                Next
            End With
        End With
        Me.hgdRecipe.Redraw = True
    End If
    
    '------------------------------------------------
    '成套方案显示(ByZT)
    If Val(Me.tvwClass.Tag) = 2 Then
        Call ShowScheme(Val(Mid(Item.Key, 2)))
    End If
    
    '------------------------------------------------
    '设置菜单和工具栏的禁止项
    If Item.ForeColor = &HFF& Then
        '已经禁止的项目不能删除
        Me.mnuEditDelete.Enabled = False
        Me.mnuEditModify.Enabled = False
        '检查部位
        'Me.mnuEditExams.Enabled = False
        '检验指标
        Me.mnuEditLabs.Enabled = False
        
        Me.mnuEditGather.Enabled = False
        '标本对照
        'Me.mnuEditSample.Enabled = False
        '收费对照
        Me.mnuEditExse.Enabled = False
        '排斥关系
        Me.mnuEditRepellent.Enabled = False
        '应用参考
'        Me.mnuEditRefer.Enabled = False
        '对应单据
        Me.mnuEditBill.Enabled = False
        
        '不能再禁止,只有启用
        Me.mnuEditStart.Enabled = (tvwClass.Tag = SK_诊疗项目 And InStr(1, mstrPrivs, "诊疗项目编辑") > 0) Or _
                            (tvwClass.Tag = SK_中药配方 And InStr(1, mstrPrivs, "中药配方编辑") > 0) Or _
                            (tvwClass.Tag = SK_成套方案 And InStr(1, mstrPrivs, "成套方案编辑") > 0)
        Me.mnuEditStop.Enabled = False
    Else
        '可以删除和修改
        Me.mnuEditDelete.Enabled = (tvwClass.Tag = SK_诊疗项目 And InStr(1, mstrPrivs, "诊疗项目编辑") > 0) Or _
                            (tvwClass.Tag = SK_中药配方 And InStr(1, mstrPrivs, "中药配方编辑") > 0) Or _
                            (tvwClass.Tag = SK_成套方案 And InStr(1, mstrPrivs, "成套方案编辑") > 0)
        Me.mnuEditModify.Enabled = Me.mnuEditDelete.Enabled
        
        '收费对照
        Me.mnuEditExse.Enabled = (InStr(1, mstrPrivs, "收费设置") > 0)
        '排斥关系
        Me.mnuEditRepellent.Enabled = True
        '应用参考
'        Me.mnuEditRefer.Enabled = (InStr(1, mstrPrivs, "参考编辑") > 0)
        '对应单据
        Me.mnuEditBill.Enabled = Me.mnuEditDelete.Enabled
        '只能停用
        Me.mnuEditStart.Enabled = False
        Me.mnuEditStop.Enabled = Me.mnuEditDelete.Enabled
        '根据类别分别判断禁止
        If Val(Me.tvwClass.Tag) = 0 And lvwItems.SelectedItem.Tag = "C" Then
            'Me.mnuEditExams.Enabled = False
            Me.mnuEditLabs.Enabled = Me.mnuEditDelete.Enabled
            Me.mnuEditGather.Enabled = Me.mnuEditDelete.Enabled
            'Me.mnuEditSample.Enabled = (InStr(1, mstrPrivs, "项目编辑") > 0)
        ElseIf Val(Me.tvwClass.Tag) = 0 And lvwItems.SelectedItem.Tag = "D" Then
            'Me.mnuEditExams.Enabled = (InStr(1, mstrPrivs, "项目编辑") > 0)
            Me.mnuEditLabs.Enabled = False
            Me.mnuEditGather.Enabled = False
            'Me.mnuEditSample.Enabled = False
        Else
            'Me.mnuEditExams.Enabled = False
            Me.mnuEditLabs.Enabled = False
            Me.mnuEditGather.Enabled = False
            'Me.mnuEditSample.Enabled = False
        End If
    End If
    
    Me.tlbThis.Buttons("Start").Enabled = Me.mnuEditStart.Enabled
    Me.tlbThis.Buttons("Stop").Enabled = Me.mnuEditStop.Enabled
    Me.tlbThis.Buttons("Delete").Enabled = Me.mnuEditDelete.Enabled
    Me.tlbThis.Buttons("Modify").Enabled = Me.mnuEditModify.Enabled
    Call zlGrdRowHeight
    Me.hgdRefer.Redraw = True
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_DblClick
End Sub

Private Sub lvwItems_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If Not lvwItems.SelectedItem Is Nothing Then
            Call PopupMenu(Me.mnuEdit, 2)
        End If
    End If
End Sub

Private Sub mnuClassAdd_Click()
    With frmClinicClass
        strTemp = Switch(Me.tvwClass.Tag = "0", 5, _
                       Me.tvwClass.Tag = "1", 4, _
                       Me.tvwClass.Tag = "2", 6)
        .lblKind.Tag = strTemp
        If Me.tvwClass.SelectedItem Is Nothing Then
            .txtParent.Tag = 0
        Else
            .txtParent.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
        End If
        .Tag = "增加"
        
        If Me.tvwClass.SelectedItem Is Nothing Then
            If .ShowMe(1, Me, "(无)", 0, 1, True) Then
                Call zlRefClasses
            End If
        Else
            If .ShowMe(1, Me, Me.tvwClass.SelectedItem.Text, Mid(Me.tvwClass.SelectedItem.Key, 2), 1, True) Then
                Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Key, 2))
            End If
        End If
    End With
End Sub

Private Sub mnuClassDel_Click()
    err = 0: On Error GoTo ErrHand
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("真的删除该分类“" & Me.tvwClass.SelectedItem.Text & "”吗", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
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
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    With frmClinicClass
        strTemp = Switch(Me.tvwClass.Tag = "0", 5, _
                       Me.tvwClass.Tag = "1", 4, _
                       Me.tvwClass.Tag = "2", 6)
        .lblKind.Tag = strTemp
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
        If Me.tvwClass.SelectedItem.Parent Is Nothing Then
            If .ShowMe(1, Me, "(无)", 0, 2, True) Then Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Key, 2))
        Else
            If .ShowMe(1, Me, Me.tvwClass.SelectedItem.Parent.Text, Mid(Me.tvwClass.SelectedItem.Parent.Key, 2), 2, True) Then Call zlRefClasses(Mid(Me.tvwClass.SelectedItem.Key, 2))
        End If
    End With
End Sub

Private Sub mnuEditAdd_Click()
    Dim blnOk As Boolean
    
    If Me.tvwClass.SelectedItem Is Nothing Then MsgBox "尚未设置分类,不能增删项目！", vbExclamation, gstrSysName: Exit Sub
    Select Case Val(Me.tvwClass.Tag)
    Case 0
        If Me.lvwItems.SelectedItem Is Nothing Then
            blnOk = frmClinicItem.ShowMe(Me, 0, Mid(Me.tvwClass.SelectedItem.Key, 2), 0)
        Else
            blnOk = frmClinicItem.ShowMe(Me, 0, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2))
        End If
    Case 1
        If Me.lvwItems.SelectedItem Is Nothing Then
            blnOk = frmMediRecipe.ShowMe(Me, 0, Mid(Me.tvwClass.SelectedItem.Key, 2), 0)
        Else
            blnOk = frmMediRecipe.ShowMe(Me, 0, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2))
        End If
    Case 2 '新增成套方案(ByZT)
        If Me.lvwItems.SelectedItem Is Nothing Then
            blnOk = frmClinicScheme.ShowMe(Me, mstrPrivs, 0, Mid(Me.tvwClass.SelectedItem.Key, 2), 0, mint范围)
        Else
            blnOk = frmClinicScheme.ShowMe(Me, mstrPrivs, 0, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2), mint范围)
        End If
    End Select
    If blnOk Then Call zlRefRecords
End Sub

Private Sub mnuEditBill_Click()
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call frmClinicBill.ShowMe(Me, Mid(Me.lvwItems.SelectedItem.Key, 2))
    Call zlRefRecords(Mid(Me.lvwItems.SelectedItem.Key, 2))
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuEditDelete_Click()
    Dim lngVItemID As Long '诊治所见项ID
    Dim rsTmp As New ADODB.Recordset
    Dim blnTrans As Boolean
    Dim blnRisTrans As Boolean
    
'    If Val(Me.tvwClass.Tag) >= 1 And Val(Me.tvwClass.Tag) <= 1 Then Exit Sub
    With Me.lvwItems
        If .SelectedItem Is Nothing Then Exit Sub
        If MsgBox("真的删除“" & .SelectedItem.Text & "”吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSql = "Select 组合项目,报告项目ID From 诊疗项目目录 A,检验报告项目 B Where A.ID=B.诊疗项目ID And A.ID=[1] "
        
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(.SelectedItem.Key, 2)))
        lngVItemID = 0
        If rsTmp.RecordCount = 1 Then If rsTmp(0) = 0 Then lngVItemID = rsTmp(1)
        
        err = 0: On Error GoTo ErrHand
                        
        '新网RIS接口；启用参数，“检查”类项目，接口部件有效的前提下
        If mblnPACSInterface = True And .SelectedItem.Tag = "D" Then
            If Not gobjRIS Is Nothing Then
                If gobjRIS.HISBasicDictTable(RISBaseItemType.ClinicItem, RISBaseItemOper.Delete, Val(Mid(.SelectedItem.Key, 2))) <> 1 Then
                    '出错时提示接口错误信息
                    If gobjRIS.LastErrorInfo <> "" Then
                        MsgBox gobjRIS.LastErrorInfo, vbInformation, gstrSysName
                    Else
                        MsgBox "调用RIS接口错误，不能继续当前操作！请与系统管理员联系", vbInformation, gstrSysName
                    End If
                    
                    Exit Sub
                End If
                blnRisTrans = True
            Else
               '接口部件无效时禁止并提示
                MsgBox "RIS接口创建失败，不能继续当前操作！可能是接口文件安装或注册不正常，请与系统管理员联系。", vbInformation, gstrSysName
                
                Exit Sub
            End If
        End If
        
        gcnOracle.BeginTrans: blnTrans = True
        
        gstrSql = "zl_诊疗项目_DELETE(" & Mid(.SelectedItem.Key, 2) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        
        If lngVItemID > 0 Then
            gstrSql = "zl_所见项目_DELETE(" & lngVItemID & ")"
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        End If
        
        gcnOracle.CommitTrans: blnTrans = False
        
        blnRisTrans = False
        
        Call .ListItems.Remove(.SelectedItem.Key)
        If .SelectedItem Is Nothing Then
            Call zlClearDetail
        Else
            Call lvwItems_ItemClick(.SelectedItem)
        End If
    End With
    Exit Sub
ErrHand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    
    Call ErrCenter
    Call SaveErrLog
    
    'Ris接口和HIS不同步时，写错误日志
    If blnRisTrans = True And Not gobjRIS Is Nothing Then
        MsgBox "HIS删除诊疗项目错误，RIS接口和HIS数据不同步，请与系统管理员联系。", vbInformation, gstrSysName
        
        On Error Resume Next
        Call gobjRIS.WriteCommLog("frmClinicLists：mnuEditDelete_Click", "HIS删除诊疗项目错误，RIS接口和HIS数据不同步", "诊疗项目ID=" & Val(Mid(lvwItems.SelectedItem.Key, 2)), 0)
    End If

End Sub

'Private Sub mnuEditExams_Click()
'    If Val(Me.tvwClass.Tag) <> 0 Then Exit Sub
'    If Me.lvwItems.SelectedItem Is Nothing Then
'        Call frmClinicPart.ShowMe(Me, True)
'    ElseIf Me.lvwItems.SelectedItem.Tag <> "D" Then
'        Call frmClinicPart.ShowMe(Me, True)
'    Else
'        If Me.lvwItems.SelectedItem.Icon = "诊疗S" Then MsgBox "停用项目，不能设置部位组合！", vbExclamation, gstrSysName
'        Call frmClinicPart.ShowMe(Me, True, Mid(Me.lvwItems.SelectedItem.Key, 2))
'    End If
'    If Not Me.lvwItems.SelectedItem Is Nothing Then
'        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
'    End If
'End Sub

Private Sub mnuEditExse_Click()
    Dim i As Integer
    Dim strIDS As String   '保存当前选中项之后的诊疗id串
    If Val(Me.tvwClass.Tag) <> 0 Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then
        Call frmClinicExse.ShowMe(Me, True)
    Else
        If Me.lvwItems.SelectedItem.Icon = "诊疗S" Then MsgBox "停用项目，不能设置收费对照！", vbExclamation, gstrSysName
        For i = Me.lvwItems.SelectedItem.Index + 1 To lvwItems.ListItems.Count
            strIDS = strIDS & Mid(Me.lvwItems.ListItems(i).Key, 2) & ","
        Next
        
        Call frmClinicExse.ShowMe(Me, True, Mid(Me.lvwItems.SelectedItem.Key, 2), strIDS)
        Call zlRefRecords(Mid(Me.lvwItems.SelectedItem.Key, 2))
    End If
    
    If Not Me.lvwItems.SelectedItem Is Nothing Then
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    End If
End Sub


Private Sub mnuEditLabs_Click()
    If Val(Me.tvwClass.Tag) <> 0 Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then
        Call frmClinicLabs.ShowMe(Me, True)
    ElseIf Me.lvwItems.SelectedItem.Tag <> "C" Then
        Call frmClinicLabs.ShowMe(Me, True)
    Else
        If Me.lvwItems.SelectedItem.Icon = "诊疗S" Then MsgBox "停用项目，不能设置检验指标！", vbExclamation, gstrSysName
        Call frmClinicLabs.ShowMe(Me, True, Mid(Me.lvwItems.SelectedItem.Key, 2))
        Call zlRefRecords(Mid(Me.lvwItems.SelectedItem.Key, 2))
    End If
    If Not Me.lvwItems.SelectedItem Is Nothing Then
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    End If
End Sub

Private Sub mnuEditModify_Click()
    Dim blnOk As Boolean
    
    If Me.tvwClass.SelectedItem Is Nothing Then MsgBox "尚未设置分类,不能增删项目！", vbExclamation, gstrSysName: Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Select Case Val(Me.tvwClass.Tag)
    Case 0
        If Me.lvwItems.SelectedItem.Icon = "诊疗S" Then MsgBox "不能对停用项目进行修改！", vbExclamation, gstrSysName
        blnOk = frmClinicItem.ShowMe(Me, 1, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2))
    Case 1
        If Me.lvwItems.SelectedItem.Icon = "方案S" Then MsgBox "不能对停用配方进行修改！", vbExclamation, gstrSysName
        blnOk = frmMediRecipe.ShowMe(Me, 1, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2))
    Case 2
        If Me.lvwItems.SelectedItem.Icon = "方案S" Then MsgBox "不能对停用方案进行修改！", vbExclamation, gstrSysName
        '修改成套方案(ByZT)
        blnOk = frmClinicScheme.ShowMe(Me, mstrPrivs, 1, Mid(Me.tvwClass.SelectedItem.Key, 2), Mid(Me.lvwItems.SelectedItem.Key, 2), mint范围)
    End Select
    If blnOk Then Call zlRefRecords(Mid(Me.lvwItems.SelectedItem.Key, 2))
    If Not Me.lvwItems.SelectedItem Is Nothing Then
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    End If
End Sub

Private Sub mnuEditRepellent_Click()
    If Val(Me.mnuEditRepellent.Tag) = 0 Then
        Call frmClinicTabu.ShowMe(Me, False)
    Else
        Call frmClinicTabu.ShowMe(Me, True)
    End If
    If Not Me.lvwItems.SelectedItem Is Nothing Then
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    End If
End Sub

Private Sub mnuEditStart_Click()
    Dim iSubItemIndex As Integer
    
    With Me.lvwItems
        If .SelectedItem Is Nothing Then Exit Sub
        If Right(.SelectedItem.Icon, 1) = "U" Then Exit Sub
        strTemp = Mid(.SelectedItem.Icon, 1, Len(.SelectedItem.Icon))
        
        If MsgBox("真的重新启用“" & .SelectedItem.Text & "”吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        gstrSql = "zl_诊疗项目_REUSE(" & Mid(.SelectedItem.Key, 2) & ")"
        err = 0: On Error GoTo ErrHand
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        If Val(Me.tvwClass.Tag) = 0 Then
            strTemp = Mid(.SelectedItem.Icon, 1, Len(.SelectedItem.Icon) - 1)
            .SelectedItem.Icon = strTemp & "U": .SelectedItem.SmallIcon = strTemp & "U"
        Else
            .SelectedItem.Icon = "方案U": .SelectedItem.SmallIcon = "方案U"
        End If
            
        '恢复启用项目显示颜色－赵彤宇
        .SelectedItem.ForeColor = .ForeColor
        For iSubItemIndex = 1 To .ColumnHeaders.Count - 1
            .SelectedItem.ListSubItems(iSubItemIndex).ForeColor = .ForeColor
        Next
    End With
    If Not Me.lvwItems.SelectedItem Is Nothing Then
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditStop_Click()
    Dim iSubItemIndex As Integer
    
    With Me.lvwItems
        If .SelectedItem Is Nothing Then Exit Sub
        If Right(.SelectedItem.Icon, 1) = "S" Then Exit Sub
        strTemp = Mid(.SelectedItem.Icon, 1, Len(.SelectedItem.Icon))
        
        If MsgBox("真的要停用“" & .SelectedItem.Text & "”吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        
        gstrSql = "zl_诊疗项目_STOP(" & Mid(.SelectedItem.Key, 2) & ")"
        err = 0: On Error GoTo ErrHand
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        If Me.mnuViewStoped.Checked = True Then
            If Val(Me.tvwClass.Tag) = 0 Then
                strTemp = Mid(.SelectedItem.Icon, 1, Len(.SelectedItem.Icon) - 1)
                .SelectedItem.Icon = strTemp & "S": .SelectedItem.SmallIcon = strTemp & "S"
            Else
                .SelectedItem.Icon = "方案S": .SelectedItem.SmallIcon = "方案S"
            End If
            
            '将停用项目显示为红色－赵彤宇
            .SelectedItem.ForeColor = &HFF&
            For iSubItemIndex = 1 To .ColumnHeaders.Count - 1
                .SelectedItem.ListSubItems(iSubItemIndex).ForeColor = &HFF&
            Next
        Else
            Call .ListItems.Remove(.SelectedItem.Key)
        End If
    End With
    If Not Me.lvwItems.SelectedItem Is Nothing Then
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuFileExcel_Click()
    Call zlRptPrint(3)
End Sub

Private Sub mnuFilePara_Click()
    Call frmClinicPara.ShowMe(Me, mstrPrivs)
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

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuhelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuViewFind_Click()
    frmClinicFind.Show , Me
End Sub

Private Sub mnuViewRefresh_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Call zlRefRecords
End Sub

Private Sub mnuViewStates_Click()
    Me.mnuViewStates.Checked = Not Me.mnuViewStates.Checked
    Me.stbThis.Visible = Me.mnuViewStates.Checked
    Form_Resize
End Sub

Private Sub mnuViewStoped_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
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
    Dim lngTop As Long, lngButtom As Long
    Dim lngNVALL As Long, lngNVBottom As Long
    
    err = 0: On Error Resume Next
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        If Not cmdKind(intCount).Visible Then
            lngNVALL = lngNVALL + 1
            If Val(Me.cmdKind(intCount).Tag) = 1 Then
                lngNVBottom = lngNVBottom + 1
            End If
        End If
    Next
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        Me.cmdKind(intCount).Left = Me.picClass.ScaleLeft + 15
        Me.cmdKind(intCount).Width = Me.picClass.ScaleWidth
        Me.cmdKind(intCount).Height = 300
        If Val(Me.cmdKind(intCount).Tag) = 0 Then
            Me.cmdKind(intCount).Top = Me.picClass.ScaleTop + lngTop
            lngTop = lngTop + IIf(cmdKind(intCount).Visible, 285, 0)
            Me.tvwClass.Top = Me.picClass.ScaleTop + lngTop
        Else
            If lngButtom = 0 Then
                lngButtom = 285 * (Me.cmdKind.UBound - intCount + 1 - lngNVBottom)
            End If
            If cmdKind(intCount).Visible Then
                Me.cmdKind(intCount).Top = Me.picClass.ScaleHeight - lngButtom
                lngButtom = lngButtom - 285
            End If
        End If
    Next
    Me.tvwClass.Left = Me.picClass.ScaleLeft + 15
    Me.tvwClass.Width = Me.picClass.ScaleWidth
    Me.tvwClass.Height = Me.picClass.ScaleHeight - 285 * (Me.cmdKind.UBound + 1 - lngNVALL) - 15
End Sub

Private Sub picHBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Me.picHBar.Top = Me.picHBar.Top + y
End Sub

Private Sub picHBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Call Form_Resize
End Sub

Private Sub picVBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Me.picVBar.Left = Me.picVBar.Left + x
End Sub

Private Sub picVBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Call Form_Resize
End Sub

Private Sub tabContent_Click(PreviousTab As Integer)
    For intCount = 0 To Me.tabContent.Tabs - 1
        If intCount = Me.tabContent.Tab Then
            Me.fraSubInfo(intCount).Visible = True
        Else
            Me.fraSubInfo(intCount).Visible = False
        End If
    Next
End Sub

Private Sub tlbThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Preview"
        Call mnuFilePreview_Click
    Case "Print"
        Call mnuFilePrint_Click
    Case "Class"
        Select Case tvwClass.Tag
            Case SK_诊疗项目
                If InStr(1, mstrPrivs, "诊疗项目编辑") > 0 Then Call PopupMenu(Me.mnuClass, 2)
            Case SK_中药配方
                If InStr(1, mstrPrivs, "中药配方编辑") > 0 Then Call PopupMenu(Me.mnuClass, 2)
            Case SK_成套方案
                If InStr(1, mstrPrivs, "成套方案编辑") > 0 Then Call PopupMenu(Me.mnuClass, 2)
        End Select
    Case "Add"
        Call mnuEditAdd_Click
    Case "Modify"
        Call mnuEditModify_Click
    Case "Delete"
        Call mnuEditDelete_Click
    Case "Start"
        Call mnuEditStart_Click
    Case "Stop"
        Call mnuEditStop_Click
    Case "Find"
        Call mnuViewFind_Click
    Case "Help"
        Call mnuHelpHelp_Click
    Case "Exit"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tlbThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Key = "add" Then
        Call mnuEditAdd_Click
    Else
        Call mnuAddcopy_Click
    End If
End Sub

Private Sub tlbThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    PopupMenu Me.mnuViewToolbar, 2
End Sub

Private Sub tvwClass_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Select Case tvwClass.Tag
            Case SK_诊疗项目
                If InStr(1, mstrPrivs, "诊疗项目编辑") > 0 Then Call PopupMenu(Me.mnuClass, 2)
            Case SK_中药配方
                If InStr(1, mstrPrivs, "中药配方编辑") > 0 Then Call PopupMenu(Me.mnuClass, 2)
            Case SK_成套方案
                If InStr(1, mstrPrivs, "成套方案编辑") > 0 Then Call PopupMenu(Me.mnuClass, 2)
        End Select
    End If
End Sub

Public Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    If Me.lvwItems.Tag = Node.Key Then Exit Sub
    Me.lvwItems.Tag = Node.Key
    Call zlRefRecords
End Sub

Private Sub zlRefClasses(Optional lngNode As Long)
    '---------------------------------------------
    '填写诊疗分类项目(此处为药品分类)并按照不同类型调整界面
    '---------------------------------------------
    
    '权限控制
    
    '调整显示界面
    Select Case Val(Me.tvwClass.Tag)
    Case 0
        Me.mnuEditAdd.Visible = True: Me.mnuEditModify.Visible = True: Me.mnuEditDelete.Visible = True: Me.mnuEditSpt1.Visible = True
'        Me.mnuEditExse.Visible = True: Me.mnuEditLabs.Visible = True: Me.mnuEditExams.Visible = True
        Me.mnuEditExse.Visible = True: Me.mnuEditLabs.Visible = False: Me.mnuEditGather.Visible = True ': Me.mnuEditSample.Visible = True
        Me.mnuEditSpt2.Visible = True: Me.mnuEditStart.Visible = True: Me.mnuEditStop.Visible = True
        Me.mnuEditSpt3.Visible = True: Me.mnuEditRepellent.Visible = True: Me.mnuEditBill.Visible = True
        Me.mnuAddcopy.Visible = True
        
        Me.tlbThis.Buttons("Split2").Visible = True
        Me.tlbThis.Buttons("Add").Visible = True: Me.tlbThis.Buttons("Modify").Visible = True: Me.tlbThis.Buttons("Delete").Visible = True
        Me.tlbThis.Buttons("Split3").Visible = True
        Me.tlbThis.Buttons("Start").Visible = True: Me.tlbThis.Buttons("Stop").Visible = True
    
    Case 1, 2
        Me.mnuEditAdd.Visible = True: Me.mnuEditModify.Visible = True: Me.mnuEditDelete.Visible = True: Me.mnuEditSpt1.Visible = False
        Me.mnuEditExse.Visible = False: Me.mnuEditLabs.Visible = False: Me.mnuEditGather.Visible = False ': Me.mnuEditSample.Visible = False
        Me.mnuEditSpt2.Visible = True: Me.mnuEditStart.Visible = True: Me.mnuEditStop.Visible = True
        Me.mnuEditSpt3.Visible = False: Me.mnuEditRepellent.Visible = False: Me.mnuEditBill.Visible = False
        
        Me.tlbThis.Buttons("Split2").Visible = True
        Me.tlbThis.Buttons("Add").Visible = True: Me.tlbThis.Buttons("Modify").Visible = True: Me.tlbThis.Buttons("Delete").Visible = True
        Me.tlbThis.Buttons("Split3").Visible = True
        Me.tlbThis.Buttons("Start").Visible = True: Me.tlbThis.Buttons("Stop").Visible = True
        Me.mnuAddcopy.Visible = False
    End Select
    
    Me.lvwItems.ListItems.Clear
    With Me.tabContent
        .TabVisible(conTab执行科室) = False
        .TabVisible(conTab收费对照) = False
        .TabVisible(conTab检验指标) = False
        .TabVisible(conTab检查部位) = False
        .TabVisible(conTab用法用量) = False
        .TabVisible(conTab配伍禁忌) = False
        .TabVisible(conTab配方组成) = False
        .TabVisible(conTab成套方案) = False
        .TabVisible(conTab应用参考) = False
    End With
    Select Case Val(Me.tvwClass.Tag)
    Case 0
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "_名称", "名称", 2500
            .Add , "_编码", "编码", 1200
            .Add , "_标本部位", "标本部位", 900
            .Add , "_计算单位", "计算单位", 900
            .Add , "_类别", "类别", 600
            .Add , "_操作类型", "操作类型", 1200
            .Add , "_执行频率", "执行频率", 900
            .Add , "_计算方式", "计算方式", 900
            .Add , "_计算规则", "计算规则", 900
            .Add , "_服务对象", "服务对象", 1200
            '.Add , "_站点", "站点", IIf(gstrNodeNo = "-", 0, 1000)
            .Add , "_院区", "院区", 600
        End With
        With Me.tabContent
            .TabVisible(conTab执行科室) = True
            .TabVisible(conTab收费对照) = True
'            .TabVisible(conTab应用参考) = True
            .Tab = conTab执行科室: Call tabContent_Click(conTab执行科室)
        End With
    Case 1
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "_名称", "名称", 2000
            .Add , "_编码", "编码", 1200
            .Add , "_说明", "说明", 3000
            '.Add , "_站点", "站点", IIf(gstrNodeNo = "-", 0, 1000)
            .Add , "_院区", "院区", 600
        End With
        With Me.tabContent
            .TabVisible(conTab配方组成) = True
            .Tab = conTab配方组成: Call tabContent_Click(conTab配方组成)
        End With
    Case 2
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "_名称", "名称", 2500
            .Add , "_编码", "编码", 1200
            .Add , "_说明", "说明", 3000
            '.Add , "_站点", "站点", IIf(gstrNodeNo = "-", 0, 1000)
            .Add , "_建档人", "建档人", "1200"
            .Add , "_建档时间", "建档时间", "2000"
            .Add , "_院区", "院区", 600
            
        End With
        With Me.tabContent
            .TabVisible(conTab成套方案) = True
            .Tab = conTab成套方案: Call tabContent_Click(conTab成套方案)
        End With
    End Select
    With Me.lvwItems
        .ColumnHeaders("_编码").Position = 1
        .SortKey = .ColumnHeaders("_编码").Index - 1: .SortOrder = lvwAscending
    End With
    Call RestoreListViewState(Me.lvwItems, Me.Name & Val(Me.tvwClass.Tag), Me.lvwItems.View)
    
    '填写分类
    err = 0: On Error GoTo ErrHand
    
    strTemp = Switch(Me.tvwClass.Tag = "0", 5, _
                   Me.tvwClass.Tag = "1", 4, _
                   Me.tvwClass.Tag = "2", 6)
    gstrSql = "select ID,上级ID,编码,名称,简码" & _
            " From 诊疗分类目录" & _
            " Where 类型 = [1] " & _
            " start with 上级ID is null" & _
            " connect by prior ID=上级ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strTemp)
    
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
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlRefRecords(Optional lngItem As Long)
    Dim iSubItemIndex As Integer
    '---------------------------------------------
    '填写项目列表
    '---------------------------------------------
    err = 0: On Error GoTo ErrHand
    
    Select Case Val(Me.tvwClass.Tag)
    Case 0
        If mnuViewShowAll.Checked = True Then
'            gstrSql = "select I.ID,I.编码,I.名称,I.标本部位,I.计算单位,I.类别 as 类别码,K.名称 as 类别,I.操作类型,I.执行频率,I.计算方式,I.计算规则," & _
'                    "        decode(I.服务对象,1,'门诊',2,'住院',3,'门诊和住院',4,'体检','不直接应用于病人') as 服务对象," & _
'                    "        nvl(I.撤档时间,to_date('3000-01-01','YYYY-MM-DD')) as 撤档时间,I.站点 " & _
'                    " from 诊疗项目目录 I,诊疗项目类别 K, " & _
'                    " (Select ID, 名称 From 诊疗分类目录 Start With 上级id = [1] Connect By Prior ID = 上级id" & _
'                    " Union ALL Select ID, 名称 From 诊疗分类目录 Where ID=[1]) B " & _
'                    " where I.类别=K.编码 And I.分类id = B.ID And (I.站点 = '" & gstrNodeNo & "' Or I.站点 is Null) "
            gstrSql = "select I.ID,I.编码,I.名称,I.标本部位,I.计算单位,I.类别 as 类别码,K.名称 as 类别,I.操作类型,I.执行频率,I.计算方式,I.计算规则," & _
                    "        decode(I.服务对象,1,'门诊',2,'住院',3,'门诊和住院',4,'体检','不直接应用于病人') as 服务对象," & _
                    "        nvl(I.撤档时间,to_date('3000-01-01','YYYY-MM-DD')) as 撤档时间,I.站点 " & _
                    " from 诊疗项目目录 I,诊疗项目类别 K, " & _
                    " (Select ID, 名称 From 诊疗分类目录 Start With 上级id = [1] Connect By Prior ID = 上级id" & _
                    " Union ALL Select ID, 名称 From 诊疗分类目录 Where ID=[1]) B " & _
                    " where I.类别=K.编码 And I.分类id = B.ID "
        Else
'            gstrSql = "select I.ID,I.编码,I.名称,I.标本部位,I.计算单位,I.类别 as 类别码,K.名称 as 类别,I.操作类型,I.执行频率,I.计算方式,I.计算规则," & _
'                    "        decode(I.服务对象,1,'门诊',2,'住院',3,'门诊和住院',4,'体检','不直接应用于病人') as 服务对象," & _
'                    "        nvl(I.撤档时间,to_date('3000-01-01','YYYY-MM-DD')) as 撤档时间,I.站点 " & _
'                    " from 诊疗项目目录 I,诊疗项目类别 K" & _
'                    " where I.类别=K.编码 and I.分类ID=[1] And (I.站点 = '" & gstrNodeNo & "' Or I.站点 is Null)"
            gstrSql = "select I.ID,I.编码,I.名称,I.标本部位,I.计算单位,I.类别 as 类别码,K.名称 as 类别,I.操作类型,I.执行频率,I.计算方式,I.计算规则," & _
                    "        decode(I.服务对象,1,'门诊',2,'住院',3,'门诊和住院',4,'体检','不直接应用于病人') as 服务对象," & _
                    "        nvl(I.撤档时间,to_date('3000-01-01','YYYY-MM-DD')) as 撤档时间,I.站点 " & _
                    " from 诊疗项目目录 I,诊疗项目类别 K" & _
                    " where I.类别=K.编码 and I.分类ID=[1] "
        End If
        If Me.mnuViewStoped.Checked = False Then
            gstrSql = gstrSql & " and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
        End If
        gstrSql = gstrSql & " order by I.编码"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Me.tvwClass.SelectedItem.Key, 2)))
        
        With rsTemp
            Me.lvwItems.ListItems.Clear
            Do While Not .EOF
                Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
                Select Case !类别码
                Case "C"
                    objItem.Icon = "检验" & IIf(Format(!撤档时间, "YYYY-MM-DD") = "3000-01-01", "U", "S")
                Case "D"
                    objItem.Icon = "检查" & IIf(Format(!撤档时间, "YYYY-MM-DD") = "3000-01-01", "U", "S")
                Case "E"
                    objItem.Icon = "处置" & IIf(Format(!撤档时间, "YYYY-MM-DD") = "3000-01-01", "U", "S")
                Case "F"
                    objItem.Icon = "手术" & IIf(Format(!撤档时间, "YYYY-MM-DD") = "3000-01-01", "U", "S")
                Case "G"
                    objItem.Icon = "麻醉" & IIf(Format(!撤档时间, "YYYY-MM-DD") = "3000-01-01", "U", "S")
                Case "H"
                    objItem.Icon = "护理" & IIf(Format(!撤档时间, "YYYY-MM-DD") = "3000-01-01", "U", "S")
                Case "I"
                    objItem.Icon = "膳食" & IIf(Format(!撤档时间, "YYYY-MM-DD") = "3000-01-01", "U", "S")
                Case "K"
                    objItem.Icon = "输血" & IIf(Format(!撤档时间, "YYYY-MM-DD") = "3000-01-01", "U", "S")
                Case "L"
                    objItem.Icon = "输氧" & IIf(Format(!撤档时间, "YYYY-MM-DD") = "3000-01-01", "U", "S")
                Case Else
                    objItem.Icon = "其他" & IIf(Format(!撤档时间, "YYYY-MM-DD") = "3000-01-01", "U", "S")
                End Select
                
                objItem.SmallIcon = objItem.Icon
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_编码").Index - 1) = !编码
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_标本部位").Index - 1) = IIf(IsNull(!标本部位), "", !标本部位)
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_计算单位").Index - 1) = IIf(IsNull(!计算单位), "", !计算单位)
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_类别").Index - 1) = !类别
                Select Case IIf(IsNull(!执行频率), 0, !执行频率)
                Case 0
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_执行频率").Index - 1) = "可选频率"
                Case 1
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_执行频率").Index - 1) = "一次性"
                Case 2
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_执行频率").Index - 1) = "持续性"
                End Select
                Select Case IIf(IsNull(!计算方式), 0, !计算方式)
                Case 0
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_计算方式").Index - 1) = "不确定"
                Case 1
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_计算方式").Index - 1) = "计量"
                Case 2
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_计算方式").Index - 1) = "计时"
                Case 3
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_计算方式").Index - 1) = "计次"
                End Select
                Select Case IIf(IsNull(!计算规则), 0, !计算规则)
                Case 0
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_计算规则").Index - 1) = "正常计算"
                Case 1
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_计算规则").Index - 1) = "取整计算"
                End Select
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_服务对象").Index - 1) = !服务对象
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_院区").Index - 1) = IIf(IsNull(!站点), "", !站点)
                Select Case !类别码
                Case "E"
                    intCount = Val(IIf(IsNull(!操作类型), 0, !操作类型))
                    strTemp = Switch(intCount = 0, "普通", _
                                    intCount = 1, "过敏试验", _
                                    intCount = 2, "给药方法(西药)", _
                                    intCount = 3, "中药煎法", _
                                    intCount = 4, "中药用(服)法", _
                                    intCount = 5, "特殊治疗", _
                                    intCount = 6, "采集方法", _
                                    intCount = 7, "配血方法", _
                                    intCount = 8, "输血途径", _
                                    intCount = 9, "输血采集")
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_操作类型").Index - 1) = strTemp
                Case "H"
                    If IIf(IsNull(!操作类型), "0", !操作类型) = "1" Then
                        objItem.SubItems(Me.lvwItems.ColumnHeaders("_操作类型").Index - 1) = "护理等级"
                    Else
                        objItem.SubItems(Me.lvwItems.ColumnHeaders("_操作类型").Index - 1) = "护理常规"
                    End If
                Case "Z"
                    intCount = Val(IIf(IsNull(!操作类型), 0, !操作类型))
                    strTemp = Switch(intCount = 0, "普通", _
                                    intCount = 1, "留观", _
                                    intCount = 2, "住院", _
                                    intCount = 3, "转科", _
                                    intCount = 4, "术后", _
                                    intCount = 5, "出院", _
                                    intCount = 6, "转院", _
                                    intCount = 7, "会诊", _
                                    intCount = 8, "抢救", _
                                    intCount = 9, "病重", _
                                    intCount = 10, "病危", _
                                    intCount = 11, "死亡", _
                                    intCount = 12, "记录入出量", _
                                    intCount = 14, "术前")
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_操作类型").Index - 1) = strTemp
                Case Else
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_操作类型").Index - 1) = IIf(IsNull(!操作类型), "", !操作类型)
                End Select
                objItem.Tag = !类别码
                If !ID = lngItem Then
                    objItem.Selected = True
                End If
                
                '将停用项目显示为红色－赵彤宇
                If Format(!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                    objItem.ForeColor = &HFF&
                    For iSubItemIndex = 1 To Me.lvwItems.ColumnHeaders.Count - 1
                        objItem.ListSubItems(iSubItemIndex).ForeColor = &HFF&
                    Next
                End If
                
                .MoveNext
            Loop
        End With
    Case 1, 2 '中药配方、成套方案
        '成套方案权限范围限制
        gstrSql = ""
        If Val(Me.tvwClass.Tag) = 2 Then
            If InStr(mstrPrivs, "全院成套方案") > 0 Then
                '有全院成套方案权限时，无限制
            ElseIf InStr(mstrPrivs, "本科成套方案") > 0 Then
                '只有本科成套方案权限时限制于本科内或自已的
                gstrSql = " And (I.人员ID=[2] Or Exists(Select 1 From 诊疗适用科室 X,部门人员 Y Where X.科室ID=Y.部门ID And X.项目ID=I.ID And Y.人员ID=[2]))"
            Else
                '都没有则只能看自已的
                gstrSql = " And I.人员ID=[2]"
            End If
        End If
        '-------------------------------
        If mnuViewShowAll.Checked = True Then
'            gstrSql = "select I.ID,I.编码,I.名称,I.标本部位,nvl(I.撤档时间,to_date('3000-01-01','YYYY-MM-DD')) as 撤档时间,I.站点 " & _
'                    " from 诊疗项目目录 I," & _
'                    " (Select ID, 名称 From 诊疗分类目录 Start With 上级id = [1] Connect By Prior ID = 上级id" & _
'                    " Union ALL Select ID, 名称 From 诊疗分类目录 Where ID=[1]) B " & _
'                    " where I.分类id = B.ID And (I.站点 = '" & gstrNodeNo & "' Or I.站点 is Null) " & gstrSql
            gstrSql = "select I.ID,I.编码,I.名称,I.标本部位,nvl(I.撤档时间,to_date('3000-01-01','YYYY-MM-DD')) as 撤档时间,I.站点,I.建档人,I.建档时间 " & _
                    " from 诊疗项目目录 I," & _
                    " (Select ID, 名称 From 诊疗分类目录 Start With 上级id = [1] Connect By Prior ID = 上级id" & _
                    " Union ALL Select ID, 名称 From 诊疗分类目录 Where ID=[1]) B " & _
                    " where I.分类id = B.ID " & gstrSql
        Else
'            gstrSql = "select I.ID,I.编码,I.名称,I.标本部位,nvl(I.撤档时间,to_date('3000-01-01','YYYY-MM-DD')) as 撤档时间,I.站点 " & _
'                    " from 诊疗项目目录 I where (I.站点 = '" & gstrNodeNo & "' Or I.站点 is Null) And I.分类ID=[1] " & gstrSql
            gstrSql = "select I.ID,I.编码,I.名称,I.标本部位,nvl(I.撤档时间,to_date('3000-01-01','YYYY-MM-DD')) as 撤档时间,I.站点,I.建档人,I.建档时间 " & _
                    " from 诊疗项目目录 I where I.分类ID=[1] " & gstrSql
        End If
        If Me.mnuViewStoped.Checked = False Then
            gstrSql = gstrSql & " and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
        End If
        gstrSql = gstrSql & " order by I.编码"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Me.tvwClass.SelectedItem.Key, 2)), UserInfo.ID)
        
        With rsTemp
            Me.lvwItems.ListItems.Clear
            Do While Not .EOF
                Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
                If Format(!撤档时间, "YYYY-MM-DD") = "3000-01-01" Then
                    objItem.Icon = "方案U": objItem.SmallIcon = "方案U"
                Else
                    objItem.Icon = "方案S": objItem.SmallIcon = "方案S"
                End If
                
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_编码").Index - 1) = !编码
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_说明").Index - 1) = IIf(IsNull(!标本部位), "", !标本部位)
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_院区").Index - 1) = IIf(IsNull(!站点), "", !站点)
                
                If Val(Me.tvwClass.Tag) = 2 Then
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_建档人").Index - 1) = IIf(IsNull(!建档人), "", !建档人)
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_建档时间").Index - 1) = IIf(IsNull(!建档时间), "", Format(!建档时间, "YYYY-MM-DD"))
                End If
                
                If !ID = lngItem Then
                    objItem.Selected = True
                End If
                
                '将停用项目显示为红色－赵彤宇
                If Format(!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                    objItem.ForeColor = &HFF&
                    For iSubItemIndex = 1 To Me.lvwItems.ColumnHeaders.Count - 1
                        objItem.ListSubItems(iSubItemIndex).ForeColor = &HFF&
                    Next
                End If
                
                .MoveNext
            Loop
        End With
    End Select

    If Me.lvwItems.ListItems.Count > 0 Then
        If Me.lvwItems.SelectedItem Is Nothing Then Me.lvwItems.ListItems(1).Selected = True
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
        err = 0: On Error Resume Next
        DoEvents: Me.lvwItems.SelectedItem.EnsureVisible
        Me.stbThis.Panels(2).Text = "该分类共有" & Me.lvwItems.ListItems.Count & "种项目"
    Else
        Call zlClearDetail
        Me.stbThis.Panels(2).Text = ""
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub InitVsfExseGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化收费对照网格
    '编制:刘兴洪
    '日期:2017-07-01 21:58:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As ListItem, str类别 As String
    Dim i As Integer, varGrade As Variant
    On Error GoTo ErrHandle
    
    Set objItem = Me.lvwItems.SelectedItem
    If Not objItem Is Nothing Then
        str类别 = objItem.SubItems(Me.lvwItems.ColumnHeaders("_类别").Index - 1)
    Else
        str类别 = ""
    End If
    
    varGrade = Split(mstrPriceGrade, ",")
    With vsfExse
        .Redraw = flexRDNone
        '.FixedRows = IIf(mblnStartPriceGrade, 2, 1)
        .Rows = .FixedRows + 1
        i = UBound(varGrade)
        i = IIf(i < 0, 0, i + IIf(mstrPriceGrade <> "", 1, 0))
        .Cols = 15 + i: .FixedCols = 1
        
        .TextMatrix(0, 0) = "":
        .TextMatrix(0, 1) = "部位":
        .TextMatrix(0, 2) = "方法"
        .TextMatrix(0, 3) = "项目名":
        .TextMatrix(0, 4) = "规格":
        .TextMatrix(0, 5) = "单位":
        .TextMatrix(0, 6) = "价格"
        .TextMatrix(0, 7) = "数量":
        .TextMatrix(0, 8) = "固定":
        .TextMatrix(0, 9) = "从项"
        .TextMatrix(0, 10) = "加收":
        .TextMatrix(0, 11) = "状态":
        .TextMatrix(0, 12) = "收费方式"
        .TextMatrix(0, 13) = "适用场合"
        .TextMatrix(0, 14) = "适用科室"
        For i = 0 To UBound(varGrade)
            .TextMatrix(0, 15 + i) = varGrade(i)
            .colData(15 + i) = "A" & i + 1
            .ColAlignment(15 + i) = flexAlignRightCenter
        Next
        For i = 0 To .Cols - 1
            .ColKey(i) = IIf(i = 0, "选择", .TextMatrix(0, i))
            .TextMatrix(.FixedRows, i) = ""
            .FixedAlignment(i) = 4
        Next
        .ColWidth(.ColIndex("选择")) = 250
        .ColWidth(.ColIndex("部位")) = 900
        .ColWidth(.ColIndex("方法")) = 1200
        .ColWidth(.ColIndex("项目名")) = 3000
        .ColWidth(.ColIndex("规格")) = 1000
        .ColWidth(.ColIndex("单位")) = 800
        .ColWidth(.ColIndex("价格")) = 1200
        .ColWidth(.ColIndex("数量")) = 1200
        .ColWidth(.ColIndex("固定")) = 600
        .ColWidth(.ColIndex("从项")) = 600
        .ColWidth(.ColIndex("加收")) = 600
        .ColWidth(.ColIndex("状态")) = 0
        .ColWidth(.ColIndex("收费方式")) = 3000
        .ColWidth(.ColIndex("适用场合")) = 850
        .ColWidth(.ColIndex("适用科室")) = 1800
        If str类别 <> "检查" Then
            .ColWidth(.ColIndex("部位")) = 0: .ColWidth(.ColIndex("方法")) = 0: .ColWidth(.ColIndex("加收")) = 0
        End If
        .ColAlignment(.ColIndex("收费方式")) = flexAlignLeftCenter
        .Redraw = flexRDBuffered
    End With
    
    

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub zlClearDetail()
    '---------------------------------------------
    '清理调整详细信息显示区域
    '---------------------------------------------
    Dim objItem As ListItem, str类别 As String
    Dim i As Integer
    If Val(Me.tvwClass.Tag) = 0 Then
        '执行科室显示
        Me.lblUseBill.Caption = "诊疗单据："
        Me.opt执行部门(0).Value = False 'True
        Me.lbl常规执行.Caption = ""
        With Me.hgd定向执行
            .Rows = .FixedRows + 1: .Cols = 2
            .TextMatrix(0, 0) = "执行科室": .TextMatrix(0, 1) = "病人科室"
            .ColWidth(0) = 1800: .ColWidth(1) = 6000
            For intCount = 0 To .Cols - 1
                .TextMatrix(.FixedRows, intCount) = "": .ColAlignmentFixed(intCount) = 4
            Next
        End With
    
        '收费对照显示
        If Val(Me.tvwClass.Tag) = 0 Then Call InitVsfExseGrid
    
        '检验指标显示
'        If Val(Me.tvwClass.Tag) = 0 Then
'            With Me.hgdLabs
'                .Rows = .FixedRows + 1: .Cols = 5: .FixedCols = 1
'                .TextMatrix(0, 0) = "": .TextMatrix(0, 1) = "检验标本": .TextMatrix(0, 2) = "报告项目": .TextMatrix(0, 3) = "类型": .TextMatrix(0, 4) = "单位"
'                .ColWidth(0) = 0: .ColWidth(1) = 0: .ColWidth(2) = 3800: .ColWidth(3) = 600: .ColWidth(4) = 1000
'                For intCount = 0 To .Cols - 1
'                    .TextMatrix(.FixedRows, intCount) = "": .ColAlignmentFixed(intCount) = 4
'                Next
'            End With
'        End If
    
        '检查部位显示
        With Me.hgdPart
            .Rows = .FixedRows + 1: .Cols = 3: .FixedCols = 1
            .TextMatrix(0, 0) = "": .TextMatrix(0, 1) = "检查项目": .TextMatrix(0, 2) = "检查部位"
            .ColWidth(0) = 250: .ColWidth(1) = 3500: .ColWidth(2) = 2000
            For intCount = 0 To .Cols - 1
                .TextMatrix(.FixedRows, intCount) = "": .ColAlignmentFixed(intCount) = 4
            Next
        End With
    End If
    
    If Val(Me.tvwClass.Tag) = 1 Then
        mbyt中药味数 = zlDatabase.GetPara(213, glngSys)
        '配方组成显示
        With Me.hgdRecipe
            .Rows = .FixedRows + 1: .Cols = mbyt中药味数 * 6: .RowHeight(0) = 0
            .GridColor = &H80000005: .BackColorBkg = &H80000005
            .MergeCells = flexMergeFree: .MergeRow(0) = True
            For intCount = 0 To .Cols - 1
                .TextMatrix(.FixedRows, intCount) = ""
                If (intCount Mod 6) = 0 Then .ColWidth(intCount) = 150
                If (intCount Mod 6) = 1 Then .ColWidth(intCount) = 0
                If (intCount Mod 6) = 2 Then .ColWidth(intCount) = 1500
                If (intCount Mod 6) = 3 Then .ColWidth(intCount) = 500
                If (intCount Mod 6) = 4 Then .ColWidth(intCount) = 200
                If (intCount Mod 6) = 5 Then .ColWidth(intCount) = 800
            Next
        End With
    End If
    
    If Val(Me.tvwClass.Tag) = 2 Then
        '成套方案显示(ByZT)
        With vsScheme
            .Rows = .FixedRows
            .Rows = .FixedRows + 1
        End With
    End If
    
    '诊疗参考显示
    With Me.hgdRefer
        .Rows = 1: .ColAlignment(0) = 1: .ColAlignment(1) = 1: .ColAlignment(2) = 1
        For intCount = 0 To .Cols - 1
            .TextMatrix(.FixedRows, intCount) = ""
        Next
    End With
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
        objPrint.Title.Text = "诊疗项目清单"
    Case 1
        objPrint.Title.Text = "中药配方清单"
    Case 2
        objPrint.Title.Text = "成套诊疗方案清单"
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

Private Sub zlGrdRowHeight()
    '---------------------------------------------
    '根据调整内容调整内容网格的行高度，以保证内容的正常显示
    '---------------------------------------------
    Dim intRow As Integer, lngColWidth As Long
    With Me.hgdRefer
        For intRow = .FixedRows To .Rows - 1
            If .TextMatrix(intRow, 1) = "" Then
                lngColWidth = .ColWidth(2)
            Else
                lngColWidth = .ColWidth(1) + .ColWidth(2)
            End If
            Me.lblScale.Width = lngColWidth - 90
            Me.lblScale.Caption = .TextMatrix(intRow, 2)
            .RowHeight(intRow) = Me.lblScale.Height + 75
        Next
    End With
End Sub

Public Sub zlLocateItem(lngClassId As Long, lngItemId As Long)
    '---------------------------------------------
    '定位到指定的诊断参考项目，在查找时使用
    '---------------------------------------------
    On Error Resume Next
    Set Me.tvwClass.SelectedItem = Me.tvwClass.Nodes("_" & lngClassId)
    Me.tvwClass.Nodes("_" & lngClassId).Selected = True
    Me.tvwClass.SelectedItem.EnsureVisible
    Call zlRefRecords
    Set Me.lvwItems.SelectedItem = Me.lvwItems.ListItems("_" & lngItemId)
    Me.lvwItems.SelectedItem.EnsureVisible
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub vsScheme_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = col内容 Then
        vsScheme.AutoSize Col
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsScheme.TextMatrix(vsScheme.FixedRows - 1, Col) & "A")
        If vsScheme.ColWidth(Col) < lngW Then
            vsScheme.ColWidth(Col) = lngW
        ElseIf vsScheme.ColWidth(Col) > vsScheme.Width * 0.5 Then
            vsScheme.ColWidth(Col) = vsScheme.Width * 0.5
        End If
    End If
End Sub

Private Sub vsScheme_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsScheme
        '擦除一并给药相关行列的边线及内容
        lngLeft = col期效: lngRight = col期效
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = col天数: lngRight = col用法
            If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        End If
        
        If Not RowIn一并给药(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '首行保留文字内容
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '底行保留下边线
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        If Between(Row, .Row, .RowSel) And Me.ActiveControl Is vsScheme Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
    Dim i As Long, blnTmp As Boolean
    With vsScheme
        If .TextMatrix(lngRow, col类别) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, col类别)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, col相关)) = Val(.TextMatrix(lngRow, col相关)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, col相关)) = Val(.TextMatrix(lngRow, col相关)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col相关)) = Val(.TextMatrix(lngRow, col相关)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col相关)) = Val(.TextMatrix(lngRow, col相关)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn一并给药 = blnTmp
    End With
End Function

Private Function ShowScheme(ByVal lng方案ID As Long) As Boolean
'功能：读取并显示数据库中的成套方案内容
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strTmp As String
    Dim str中药 As String, str煎法 As String
    Dim str麻醉 As String, Str标本 As String
    Dim i As Long, j As Long

    On Error GoTo errH

    strSql = "Select A.序号,A.相关序号,A.期效,A.诊疗项目ID,A.医嘱内容,A.天数," & _
             " A.单次用量,A.执行频次,A.医生嘱托,Nvl(C.名称,Decode(Nvl(A.执行性质,0),0,'<叮嘱>',5,'-')) as 执行科室," & _
             " A.执行性质,A.执行标记,A.时间方案,Nvl(B.类别,'*') as 类别,Nvl(D.名称||Decode(D.规格,NULL,NULL,' '||D.规格),B.名称) as 名称," & _
             " B.计算单位,A.标本部位,A.检查方法,A.总给予量,D.计算单位 as 总量单位,D.ID as 收费细目ID," & _
             " Nvl(B.撤档时间,To_Date('3000-01-01','YYYY-MM-DD')) As 撤档时间" & _
             " From 诊疗项目组合 A,诊疗项目目录 B,部门表 C,收费项目目录 D" & _
             " Where A.诊疗项目ID=B.ID(+) And A.执行科室ID=C.ID(+)" & _
             " And A.收费细目id=D.ID(+) And A.诊疗组合ID=[1] " & _
             " Order by A.序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng方案ID)

    With vsScheme
        .Redraw = flexRDNone
        .Rows = .FixedRows    '清除表格内容
        If rsTmp.EOF Then
            .Rows = .FixedRows + 1
        Else
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, col期效) = IIf(NVL(rsTmp!期效, 0) = 0, "长期", "临时")
                .TextMatrix(i, col内容) = NVL(rsTmp!医嘱内容, NVL(rsTmp!名称))
                .TextMatrix(i, col标本部位) = NVL(rsTmp!标本部位)    '检验标本
                .TextMatrix(i, col检查方法) = NVL(rsTmp!检查方法)
                .TextMatrix(i, col单量) = FormatEx(NVL(rsTmp!单次用量), 4)
                If Not IsNull(rsTmp!单次用量) Then
                    If rsTmp!类别 = "4" Then
                        .TextMatrix(i, col单位) = NVL(rsTmp!总量单位)
                    Else
                        .TextMatrix(i, col单位) = NVL(rsTmp!计算单位)
                    End If
                End If
                If .TextMatrix(i, col期效) = "临时" Then
                    If Not IsNull(rsTmp!总给予量) Then
                        .TextMatrix(i, col总量) = FormatEx(NVL(rsTmp!总给予量), 4)
                        If Not IsNull(rsTmp!总量单位) Then
                            .TextMatrix(i, col总量单位) = NVL(rsTmp!总量单位)
                        ElseIf InStr(",4,5,6,7,", rsTmp!类别) = 0 Then
                            .TextMatrix(i, col总量单位) = NVL(rsTmp!计算单位)
                        End If
                    End If
                End If
                .TextMatrix(i, col天数) = NVL(rsTmp!天数)
                .TextMatrix(i, col频次) = NVL(rsTmp!执行频次)
                .TextMatrix(i, col嘱托) = NVL(rsTmp!医生嘱托)
                .TextMatrix(i, col执行时间) = NVL(rsTmp!时间方案)
                .TextMatrix(i, col执行科室) = NVL(rsTmp!执行科室)
                .Cell(flexcpData, i, col执行性质) = NVL(rsTmp!执行性质, 0)
                .TextMatrix(i, col序号) = rsTmp!序号
                .TextMatrix(i, col相关) = NVL(rsTmp!相关序号)
                .TextMatrix(i, col项目ID) = NVL(rsTmp!诊疗项目id)
                .TextMatrix(i, col收费细目ID) = NVL(rsTmp!收费细目ID)
                .TextMatrix(i, col类别) = rsTmp!类别
                .TextMatrix(i, col执行标记) = NVL(rsTmp!执行标记)
                .TextMatrix(i, col停用) = IIf(Format(rsTmp!撤档时间, "YYYY-MM-DD") <> "3000-01-01", "√", "")
                If Format(rsTmp!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &HFF&
                End If
                rsTmp.MoveNext
            Next

            '再处理一些附加行的隐藏,及相关内容的显示
            For i = 1 To .Rows - 1
                '给药途径
                If .TextMatrix(i, col类别) = "E" And Val(.TextMatrix(i, col相关)) = 0 _
                   And Val(.TextMatrix(i - 1, col相关)) = Val(.TextMatrix(i, col序号)) _
                   And InStr(",5,6,", .TextMatrix(i - 1, col类别)) > 0 Then
                    .RowHidden(i) = True
                    '显示给药途径
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, col相关)) = Val(.TextMatrix(i, col序号)) Then
                            .TextMatrix(j, col用法) = .TextMatrix(i, col内容)

                            '显示成药的执行性质
                            If Val(.Cell(flexcpData, j, col执行性质)) = 5 And Val(.Cell(flexcpData, i, col执行性质)) <> 5 Then
                                .TextMatrix(j, col执行性质) = IIf(Val(.TextMatrix(j, col执行标记)) = 2, "不取药", "自备药")
                            ElseIf Val(.Cell(flexcpData, j, col执行性质)) <> 5 And Val(.Cell(flexcpData, i, col执行性质)) = 5 Then
                                .TextMatrix(j, col执行性质) = "离院带药"
                            Else
                                .TextMatrix(j, col执行性质) = IIf(Val(.TextMatrix(j, col执行标记)) = 1, "自取药", "正常")
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If

                '输血途径
                If .TextMatrix(i, col类别) = "E" And .TextMatrix(i - 1, col类别) = "K" _
                   And Val(.TextMatrix(i, col相关)) = Val(.TextMatrix(i - 1, col序号)) Then
                    .RowHidden(i) = True
                    .TextMatrix(i - 1, col用法) = .TextMatrix(i, col内容)
                    .TextMatrix(i - 1, col内容) = .TextMatrix(i - 1, col内容) & "(" & .TextMatrix(i, col内容) & ")"
                End If

                '中药配方和检验组合
                If .TextMatrix(i, col类别) = "E" And Val(.TextMatrix(i, col相关)) = 0 _
                   And Val(.TextMatrix(i - 1, col相关)) = Val(.TextMatrix(i, col序号)) _
                   And InStr(",7,E,C,", .TextMatrix(i - 1, col类别)) > 0 Then

                    str中药 = "": str煎法 = "": Str标本 = "": strTmp = ""
                    j = .FindRow(CStr(Val(.TextMatrix(i, col序号))), , col相关)

                    '中药及检验的执行科室
                    .TextMatrix(i, col执行科室) = .TextMatrix(j, col执行科室)

                    '显示中药配方执行性质:以药品为准判断
                    If .TextMatrix(i - 1, col类别) <> "C" Then
                        If Val(.Cell(flexcpData, j, col执行性质)) = 5 And Val(.Cell(flexcpData, i, col执行性质)) <> 5 Then
                            .TextMatrix(i, col执行性质) = IIf(Val(.TextMatrix(i, col执行标记)) = 2, "不取药", "自备药")
                        ElseIf Val(.Cell(flexcpData, j, col执行性质)) <> 5 And Val(.Cell(flexcpData, i, col执行性质)) = 5 Then
                            .TextMatrix(i, col执行性质) = "离院带药"
                        Else
                            .TextMatrix(i, col执行性质) = IIf(Val(.TextMatrix(i, col执行标记)) = 1, "自取药", "正常")
                        End If
                    End If

                    For j = j To i - 1
                        .RowHidden(j) = j <> i
                        If .TextMatrix(j, col类别) = "7" Then
                            str中药 = str中药 & "," & RTrim(.TextMatrix(j, col内容) & _
                                                        " " & .TextMatrix(j, col单量) & .TextMatrix(j, col单位) & _
                                                        " " & .TextMatrix(j, col嘱托))
                        ElseIf .TextMatrix(j, col类别) = "C" Then
                            strTmp = strTmp & "," & .TextMatrix(j, col内容)
                            Str标本 = .TextMatrix(j, col标本部位)    '取第一个检验项目的标本
                        ElseIf .TextMatrix(j, col类别) = "E" And Val(.TextMatrix(j, col相关)) <> 0 Then
                            str煎法 = .TextMatrix(j, col内容) & .TextMatrix(j, col标本部位)
                        End If
                    Next

                    .TextMatrix(i, col用法) = .TextMatrix(i, col内容)    '显示中药用法或检验采集方法

                    If .TextMatrix(i - 1, col类别) = "C" Then
                        .TextMatrix(i, col内容) = Mid(strTmp, 2) & IIf(Str标本 <> "", "(" & Str标本 & ")", "")
                    Else
                        .TextMatrix(i, col内容) = "中药配方," & .TextMatrix(i, col频次) & "," & _
                                                str煎法 & "," & .TextMatrix(i, col内容) & ":" & Mid(str中药, 2)
                        .TextMatrix(i, col总量单位) = "付"
                    End If
                End If

                '检查组合
                If .TextMatrix(i, col类别) = "D" And Val(.TextMatrix(i, col相关)) = 0 Then
                    Str标本 = "": str煎法 = "": strTmp = ""
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, col相关)) = Val(.TextMatrix(i, col序号)) Then
                            .RowHidden(j) = True
                            If .TextMatrix(j, col标本部位) <> "" _
                               And Val(.TextMatrix(j, col项目ID)) = Val(.TextMatrix(i, col项目ID)) Then    '相同的项目ID才是新方式
                                If .TextMatrix(j, col标本部位) <> strTmp And strTmp <> "" Then
                                    Str标本 = Str标本 & "," & strTmp & IIf(str煎法 <> "", "(" & Mid(str煎法, 2) & ")", "")
                                    str煎法 = ""
                                End If
                                If .TextMatrix(j, col检查方法) <> "" Then
                                    str煎法 = str煎法 & "," & .TextMatrix(j, col检查方法)
                                End If

                                strTmp = .TextMatrix(j, col标本部位)
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    If strTmp <> "" Then
                        Str标本 = Str标本 & "," & strTmp & IIf(str煎法 <> "", "(" & Mid(str煎法, 2) & ")", "")
                    End If
                    If Str标本 <> "" Then    '以前的检查方式时不显示详细医嘱内容
                        .TextMatrix(i, col内容) = .TextMatrix(i, col内容) & ":" & Mid(Str标本, 2)
                    End If
                End If

                '手术项目
                If .TextMatrix(i, col类别) = "F" And Val(.TextMatrix(i, col相关)) = 0 Then
                    strTmp = "": str麻醉 = ""
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, col相关)) = Val(.TextMatrix(i, col序号)) Then
                            .RowHidden(j) = True
                            If .TextMatrix(j, col类别) = "F" Then
                                strTmp = strTmp & "," & .TextMatrix(j, col内容)
                            ElseIf .TextMatrix(j, col类别) = "G" Then
                                str麻醉 = .TextMatrix(j, col内容)
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    If strTmp <> "" Or str麻醉 <> "" Then
                        If str麻醉 <> "" Then
                            .TextMatrix(i, col内容) = "在 " & str麻醉 & " 下行 " & .TextMatrix(i, col内容)
                        Else
                            .TextMatrix(i, col内容) = "行 " & .TextMatrix(i, col内容)
                        End If
                        If strTmp <> "" Then
                            .TextMatrix(i, col内容) = .TextMatrix(i, col内容) & " 及 " & Mid(strTmp, 2)
                        End If
                    End If
                End If
            Next
        End If
        .Row = .FixedRows: .Col = .FixedCols
        .AutoSize col内容
        .Redraw = flexRDDirect
    End With
    ShowScheme = True
    Exit Function
errH:
    vsScheme.Redraw = flexRDDirect
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub
